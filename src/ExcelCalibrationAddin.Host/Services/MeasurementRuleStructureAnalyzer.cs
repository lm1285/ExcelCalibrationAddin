using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleStructureAnalyzer
    {
        private static readonly Regex CellReferenceRegex = new Regex(@"\$?[A-Z]{1,3}\$?\d+", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public void Apply(WorkbookSnapshot snapshot, IReadOnlyList<MeasurementRule> rules)
        {
            if (snapshot == null || rules == null)
            {
                return;
            }

            foreach (var rule in rules.Where(item => item != null))
            {
                var sheet = ResolveSheet(snapshot, rule);
                if (sheet == null)
                {
                    continue;
                }

                // Formula text and cached results belong to the current
                // workbook.  If the located cell currently contains a formula,
                // always rebuild the formula descriptor so edits at the same
                // address invalidate the old interpretation.  Synthetic/unit
                // test rules without a snapshot formula retain their supplied
                // descriptor for backwards compatibility.
                var currentFormula = ResolveErrorFormula(sheet, rule);
                if (currentFormula.HasFormula)
                {
                    rule.ErrorFormula = currentFormula;
                    ApplySupplementalFormulaInfo(sheet, rule, rule.ErrorFormula);
                }
                else if (!HasTemplateFormulaInfo(rule.ErrorFormula))
                {
                    rule.ErrorFormula = HasBaseFormulaInfo(rule.ErrorFormula)
                        ? rule.ErrorFormula
                        : ResolveErrorFormula(sheet, rule);
                    ApplySupplementalFormulaInfo(sheet, rule, rule.ErrorFormula);
                }
                else
                {
                    // Even when the primary formula is absent, clear stale
                    // supplemental formulas that may have been removed.
                    RefreshSupplementalFormulaInfo(sheet, rule, rule.ErrorFormula);
                }

                RefreshFormulaClassification(sheet, rule, rule.ErrorFormula);
                AlignStandardValueSourceToFormula(sheet, rule, rule.ErrorFormula);
                RefreshFormulaDependencies(snapshot, rule, rule.ErrorFormula);
                ApplyAdditionalJudgementConstraints(snapshot, sheet, rule);
            }
        }

        private static void ApplyAdditionalJudgementConstraints(
            WorkbookSnapshot snapshot,
            SheetSnapshot sheet,
            MeasurementRule rule)
        {
            if (rule == null || (rule.AdditionalJudgementConstraints ?? new List<MeasurementJudgementConstraint>()).Count == 0)
            {
                return;
            }

            var original = MeasurementJudgementConstraintHelper.Capture(rule, null);
            var resolvedConstraints = new List<MeasurementJudgementConstraint>();
            foreach (var constraint in rule.AdditionalJudgementConstraints)
            {
                if (constraint == null)
                {
                    continue;
                }

                MeasurementJudgementConstraintHelper.Overlay(rule, constraint);
                if (!HasTemplateFormulaInfo(rule.ErrorFormula))
                {
                    rule.ErrorFormula = HasBaseFormulaInfo(rule.ErrorFormula)
                        ? rule.ErrorFormula
                        : ResolveErrorFormula(sheet, rule);
                    ApplySupplementalFormulaInfo(sheet, rule, rule.ErrorFormula);
                }

                RefreshFormulaClassification(sheet, rule, rule.ErrorFormula);
                RefreshFormulaDependencies(snapshot, rule, rule.ErrorFormula);
                resolvedConstraints.Add(MeasurementJudgementConstraintHelper.Capture(rule, constraint));
            }

            MeasurementJudgementConstraintHelper.Overlay(rule, original);
            rule.AdditionalJudgementConstraints = resolvedConstraints;
        }

        private static void RefreshFormulaDependencies(
            WorkbookSnapshot snapshot,
            MeasurementRule rule,
            ErrorFormulaInfo info)
        {
            if (info == null)
            {
                return;
            }

            var fallbackSheetName = rule?.TargetRange?.SheetName ??
                rule?.ErrorSource?.Range?.SheetName ?? string.Empty;
            info.DependencyRanges = new[]
                {
                    info.Formula,
                    info.AverageFormula,
                    info.TechnicalRequirementFormula,
                    info.UncertaintyFormula,
                    info.ResultFormula
                }
                .Where(formula => !string.IsNullOrWhiteSpace(formula))
                .SelectMany(formula => TemplateFormulaParser.ExtractReferencedRanges(formula, fallbackSheetName))
                .Concat(new[]
                {
                    info.Formula,
                    info.AverageFormula,
                    info.TechnicalRequirementFormula,
                    info.UncertaintyFormula,
                    info.ResultFormula
                }
                .Where(formula => !string.IsNullOrWhiteSpace(formula))
                .SelectMany(formula => TemplateFormulaParser.ExtractNamedRangeReferences(formula, snapshot?.NamedRanges)))
                .Where(range => range != null && !string.IsNullOrWhiteSpace(range.SheetName))
                .GroupBy(range => string.Join(":", new[]
                {
                    range.SheetName.ToUpperInvariant(),
                    range.StartRow.ToString(CultureInfo.InvariantCulture),
                    range.StartColumn.ToString(CultureInfo.InvariantCulture),
                    range.EndRow.ToString(CultureInfo.InvariantCulture),
                    range.EndColumn.ToString(CultureInfo.InvariantCulture)
                }))
                .Select(group => group.First())
                .ToList();
            info.UnresolvedDependencies = new[]
                {
                    info.Formula,
                    info.AverageFormula,
                    info.TechnicalRequirementFormula,
                    info.UncertaintyFormula,
                    info.ResultFormula
                }
                .Where(formula => !string.IsNullOrWhiteSpace(formula))
                .SelectMany(formula => TemplateFormulaParser.FindUnresolvedDependencies(
                    formula,
                    (snapshot?.NamedRanges ?? new List<NamedRangeDefinition>()).Select(item => item?.Name)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void RefreshFormulaClassification(
            SheetSnapshot sheet,
            MeasurementRule rule,
            ErrorFormulaInfo info)
        {
            if (sheet == null || rule == null || string.IsNullOrWhiteSpace(info?.Formula))
            {
                return;
            }

            string classificationFormula;
            var usedActiveBranch = TemplateFormulaParser.TryResolveActiveValueExpression(
                sheet,
                info.Formula,
                out classificationFormula);
            if (!usedActiveBranch)
            {
                classificationFormula = info.Formula;
            }

            var references = ExtractReferences(info.Formula);
            info.ReferencesMeasurement = references.Any(reference => RangeContains(rule.TargetRange, reference.Row, reference.Column));
            info.ReferencesStandardValue = references.Any(reference => RangeContains(rule.StandardValueSource?.Range, reference.Row, reference.Column));
            info.ReferencesAverage = references.Any(reference => RangeContains(rule.AverageSource?.Range, reference.Row, reference.Column));
            info.FormulaMultipliesBy100 = FormulaMultipliesBy100(classificationFormula);
            info.FormulaDividesByReferenceRange = FormulaDividesByRange(classificationFormula, rule.RangeSource?.Range);
            info.Scale = ResolveFormulaScale(sheet, rule, info, classificationFormula, usedActiveBranch);
        }

        private static void AlignStandardValueSourceToFormula(
            SheetSnapshot sheet,
            MeasurementRule rule,
            ErrorFormulaInfo info)
        {
            if (sheet == null || rule?.TargetRange == null || string.IsNullOrWhiteSpace(info?.Formula))
            {
                return;
            }

            var references = ExtractReferences(info.Formula);
            if (references.Any(reference =>
                RangeContains(rule.StandardValueSource?.Range, reference.Row, reference.Column)))
            {
                info.ReferencesStandardValue = true;
                return;
            }

            // Explicitly captured standard/setpoint positions are authoritative.
            // Formula analysis must not reinterpret the setpoint column as the
            // legacy standard-value column merely because the formula references it.
            if (rule.SetpointSource?.Range != null)
            {
                return;
            }

            var candidateColumns = references
                .Where(reference => reference.Row >= rule.TargetRange.StartRow && reference.Row <= rule.TargetRange.EndRow)
                .Where(reference => !RangeContains(rule.TargetRange, reference.Row, reference.Column))
                .Where(reference => !RangeContains(rule.AverageSource?.Range, reference.Row, reference.Column))
                .Where(reference => !RangeContains(rule.ErrorSource?.Range, reference.Row, reference.Column))
                .Where(reference => !RangeContains(rule.MpeSource?.Range, reference.Row, reference.Column))
                .Where(reference => !RangeContains(rule.RangeSource?.Range, reference.Row, reference.Column))
                .Where(reference => !RangeContains(rule.UncertaintySource?.Range, reference.Row, reference.Column))
                .Where(reference => !RangeContains(rule.ResultSource?.Range, reference.Row, reference.Column))
                .Select(reference => reference.Column)
                .Distinct()
                .ToList();
            if (candidateColumns.Count != 1)
            {
                return;
            }

            var logicalRanges = new List<CellRange>();
            for (var row = rule.TargetRange.StartRow; row <= rule.TargetRange.EndRow; row++)
            {
                var column = candidateColumns[0];
                var logical = GetLogicalCells(sheet, new CellRange
                {
                    SheetName = rule.TargetRange.SheetName,
                    StartRow = row,
                    EndRow = row,
                    StartColumn = column,
                    EndColumn = column
                }).FirstOrDefault(item => row >= item.Range.StartRow && row <= item.Range.EndRow);
                if (logical == null)
                {
                    return;
                }

                logicalRanges.Add(logical.Range);
            }

            var inferredRange = new CellRange
            {
                SheetName = rule.TargetRange.SheetName,
                StartRow = logicalRanges.Min(range => range.StartRow),
                EndRow = logicalRanges.Max(range => range.EndRow),
                StartColumn = logicalRanges.Min(range => range.StartColumn),
                EndColumn = logicalRanges.Max(range => range.EndColumn)
            };
            if (rule.StandardValueSource == null)
            {
                rule.StandardValueSource = new ParameterSource { Name = "标准值" };
            }

            rule.StandardValueSource.Range = inferredRange;
            info.ReferencesStandardValue = true;
        }

        private static bool HasTemplateFormulaInfo(ErrorFormulaInfo info)
        {
            return info != null &&
                HasBaseFormulaInfo(info) &&
                !string.IsNullOrWhiteSpace(info.TechnicalRequirementFormula) &&
                !string.IsNullOrWhiteSpace(info.ResultFormula);
        }

        private static bool HasBaseFormulaInfo(ErrorFormulaInfo info)
        {
            return info != null &&
                info.HasFormula &&
                (info.ReferencesMeasurement ||
                 (info.ReferencesAverage && info.AverageFormulaResolved));
        }

        private static SheetSnapshot ResolveSheet(WorkbookSnapshot snapshot, MeasurementRule rule)
        {
            var sheetName = rule?.TargetRange?.SheetName ??
                            rule?.SetpointSource?.Range?.SheetName ??
                            rule?.StandardValueSource?.Range?.SheetName ??
                            rule?.ErrorSource?.Range?.SheetName;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return null;
            }

            return snapshot.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        }

        private static ErrorFormulaInfo ResolveErrorFormula(SheetSnapshot sheet, MeasurementRule rule)
        {
            var info = new ErrorFormulaInfo();
            if (sheet == null || rule?.ErrorSource?.Range == null)
            {
                return info;
            }

            var formulaCell = GetLogicalCells(sheet, rule.ErrorSource.Range)
                .Select(item => item.Anchor)
                .FirstOrDefault(item => !string.IsNullOrWhiteSpace(item?.Formula));
            if (formulaCell == null)
            {
                return info;
            }

            info.HasFormula = true;
            info.Formula = formulaCell.Formula ?? string.Empty;
            var references = ExtractReferences(info.Formula);
            info.ReferencesMeasurement = references.Any(reference => RangeContains(rule.TargetRange, reference.Row, reference.Column));
            info.ReferencesStandardValue = references.Any(reference => RangeContains(rule.StandardValueSource?.Range, reference.Row, reference.Column));
            info.ReferencesAverage = references.Any(reference => RangeContains(rule.AverageSource?.Range, reference.Row, reference.Column));
            string classificationFormula;
            var usedActiveBranch = TemplateFormulaParser.TryResolveActiveValueExpression(
                sheet,
                info.Formula,
                out classificationFormula);
            if (!usedActiveBranch)
            {
                classificationFormula = info.Formula;
            }

            info.FormulaMultipliesBy100 = FormulaMultipliesBy100(classificationFormula);
            info.FormulaDividesByReferenceRange = FormulaDividesByRange(classificationFormula, rule.RangeSource?.Range);
            info.Scale = ResolveFormulaScale(sheet, rule, info, classificationFormula, usedActiveBranch);

            if (info.ReferencesAverage && rule.AverageSource?.Range != null)
            {
                var averageCell = GetLogicalCells(sheet, rule.AverageSource.Range)
                    .Select(item => item.Anchor)
                    .FirstOrDefault(item => !string.IsNullOrWhiteSpace(item?.Formula));
                info.AverageFormula = averageCell?.Formula ?? string.Empty;
                var averageReferences = ExtractReferences(info.AverageFormula);
                info.AverageFormulaResolved = averageReferences.Any(reference => RangeContains(rule.TargetRange, reference.Row, reference.Column)) ||
                    info.AverageFormula.IndexOf("AVERAGE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    info.AverageFormula.IndexOf("AVG", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            return info;
        }

        private static void ApplySupplementalFormulaInfo(SheetSnapshot sheet, MeasurementRule rule, ErrorFormulaInfo info)
        {
            if (info == null)
            {
                return;
            }

            RefreshSupplementalFormulaInfo(sheet, rule, info);
        }

        private static void RefreshSupplementalFormulaInfo(SheetSnapshot sheet, MeasurementRule rule, ErrorFormulaInfo info)
        {
            if (info == null)
            {
                return;
            }

            info.TechnicalRequirementFormula = string.Empty;
            info.TechnicalRequirementFormulaResolved = false;
            info.UncertaintyFormula = string.Empty;
            info.UncertaintyFormulaResolved = false;
            info.ResultFormula = string.Empty;
            info.ResultFormulaResolved = false;

            var technicalRequirementFormula = ResolveFirstFormula(sheet, rule?.MpeSource?.Range);
            if (!string.IsNullOrWhiteSpace(technicalRequirementFormula))
            {
                info.TechnicalRequirementFormula = technicalRequirementFormula;
                info.TechnicalRequirementFormulaResolved = true;
            }

            var uncertaintyFormula = ResolveFirstFormula(sheet, rule?.UncertaintySource?.Range);
            if (!string.IsNullOrWhiteSpace(uncertaintyFormula))
            {
                info.UncertaintyFormula = uncertaintyFormula;
                info.UncertaintyFormulaResolved = true;
            }

            var resultFormula = ResolveFirstFormula(sheet, rule?.ResultSource?.Range);
            if (!string.IsNullOrWhiteSpace(resultFormula))
            {
                info.ResultFormula = resultFormula;
                info.ResultFormulaResolved = true;
            }
        }

        private static string ResolveFirstFormula(SheetSnapshot sheet, CellRange range)
        {
            return GetLogicalCells(sheet, range)
                .Select(item => item.Anchor)
                .Select(cell => cell?.Formula)
                .FirstOrDefault(formula => !string.IsNullOrWhiteSpace(formula)) ?? string.Empty;
        }

        private static ErrorFormulaScale ResolveFormulaScale(
            SheetSnapshot sheet,
            MeasurementRule rule,
            ErrorFormulaInfo info,
            string classificationFormula,
            bool usedActiveBranch)
        {
            if (usedActiveBranch)
            {
                if (info.FormulaDividesByReferenceRange)
                {
                    return ErrorFormulaScale.RelativeToReferenceRange;
                }

                if (LooksLikeRelativeFormula(classificationFormula))
                {
                    return ErrorFormulaScale.RelativeToStandardValue;
                }
            }

            // The requirement cell is the authoritative signal for conditional
            // formulas whose active branch changes both tolerance and unit.
            var requirementScale = ResolveScaleFromContext(CollectRangeContext(sheet, rule?.MpeSource?.Range));
            if (requirementScale.HasValue)
            {
                return requirementScale.Value;
            }

            var errorScale = ResolveScaleFromContext(CollectRangeContext(sheet, rule?.ErrorSource?.Range));
            if (errorScale.HasValue)
            {
                return errorScale.Value;
            }

            if (info.FormulaDividesByReferenceRange)
            {
                return ErrorFormulaScale.RelativeToReferenceRange;
            }

            if (LooksLikeRelativeFormula(classificationFormula))
            {
                return ErrorFormulaScale.RelativeToStandardValue;
            }

            return ErrorFormulaScale.Absolute;
        }

        private static ErrorFormulaScale? ResolveScaleFromContext(string context)
        {
            if (ContainsReferencedSignal(context))
            {
                return ErrorFormulaScale.RelativeToReferenceRange;
            }

            if (ContainsPercentToken(context))
            {
                return ErrorFormulaScale.RelativeToStandardValue;
            }

            return ContainsAbsoluteErrorSignal(context)
                ? ErrorFormulaScale.Absolute
                : (ErrorFormulaScale?)null;
        }

        private static string CollectRangeContext(SheetSnapshot sheet, CellRange range)
        {
            var values = new List<string>();
            AddRangeContext(values, sheet, range);
            return string.Join(" ", values);
        }

        private static void AddRangeContext(ICollection<string> values, SheetSnapshot sheet, CellRange range)
        {
            if (sheet == null || range == null)
            {
                return;
            }

            var headerStartRow = Math.Max(1, range.StartRow - 4);
            foreach (var cell in GetLogicalCells(sheet, new CellRange
            {
                SheetName = range.SheetName,
                StartRow = headerStartRow,
                EndRow = Math.Max(headerStartRow, range.StartRow - 1),
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            }).Concat(GetLogicalCells(sheet, range)))
            {
                AddContextValue(values, cell.Anchor?.Text);
                AddContextValue(values, cell.Anchor?.NumberFormat);
            }
        }

        private static void AddContextValue(ICollection<string> values, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            var trimmed = value.Trim();
            if (trimmed.Length > 0 && !values.Contains(trimmed))
            {
                values.Add(trimmed);
            }
        }

        private static bool LooksLikeRelativeFormula(string formula)
        {
            var normalized = NormalizeFormula(formula);
            return normalized.Contains("/") && normalized.Contains("*100");
        }

        private static bool FormulaMultipliesBy100(string formula)
        {
            return NormalizeFormula(formula).Contains("*100");
        }

        private static bool FormulaDividesByRange(string formula, CellRange range)
        {
            if (string.IsNullOrWhiteSpace(formula) || range == null)
            {
                return false;
            }

            foreach (var denominator in ExtractFormulaDenominators(formula))
            {
                if (ExtractReferences(denominator)
                    .Any(reference => RangeContains(range, reference.Row, reference.Column)))
                {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<string> ExtractFormulaDenominators(string formula)
        {
            var text = formula ?? string.Empty;
            var inString = false;
            for (var index = 0; index < text.Length; index++)
            {
                if (text[index] == '"')
                {
                    inString = !inString;
                    continue;
                }

                if (inString || text[index] != '/')
                {
                    continue;
                }

                var start = index + 1;
                while (start < text.Length && char.IsWhiteSpace(text[start]))
                {
                    start++;
                }

                if (start >= text.Length)
                {
                    continue;
                }

                var end = start;
                if (text[start] == '(')
                {
                    var depth = 0;
                    for (; end < text.Length; end++)
                    {
                        if (text[end] == '(')
                        {
                            depth++;
                        }
                        else if (text[end] == ')' && --depth == 0)
                        {
                            end++;
                            break;
                        }
                    }
                }
                else
                {
                    while (end < text.Length && !IsFormulaOperandBoundary(text[end]))
                    {
                        end++;
                    }
                }

                if (end > start)
                {
                    yield return text.Substring(start, end - start);
                }
            }
        }

        private static bool IsFormulaOperandBoundary(char value)
        {
            return char.IsWhiteSpace(value) ||
                value == '+' ||
                value == '-' ||
                value == '*' ||
                value == '/' ||
                value == ',' ||
                value == ';' ||
                value == '<' ||
                value == '>' ||
                value == '=' ||
                value == ')';
        }

        private static bool ContainsPercentToken(string text)
        {
            return NormalizeSignal(text).Contains("%");
        }

        private static bool ContainsAbsoluteErrorSignal(string text)
        {
            var normalized = NormalizeSignal(text)
                .Replace("Μ", "U");
            return normalized.Contains("绝对误差") ||
                normalized.Contains("UMOL/MOL") ||
                normalized.Contains("µMOL/MOL") ||
                normalized.Contains("MOL/MOL");
        }

        private static bool ContainsReferencedSignal(string text)
        {
            var normalized = NormalizeSignal(text);
            return normalized.Contains("%FS") ||
                normalized.Contains("FS%") ||
                normalized.Contains("FULLSCALE") ||
                normalized.Contains("FULL SCALE") ||
                normalized.Contains("SPAN") ||
                normalized.Contains("引用误差") ||
                normalized.Contains("满量程误差") ||
                normalized.Contains("全量程误差");
        }

        private static string NormalizeFormula(string formula)
        {
            return NormalizeSignal(formula).Replace(" ", string.Empty);
        }

        private static string NormalizeSignal(string text)
        {
            return (text ?? string.Empty)
                .Trim()
                .Replace("\uFF05", "%")
                .Replace("\uFF08", "(")
                .Replace("\uFF09", ")")
                .Replace("\uFF0D", "-")
                .Replace("\u2013", "-")
                .Replace("\u2014", "-")
                .ToUpperInvariant();
        }

    }
}
