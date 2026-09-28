using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Recognition;
using ExcelCalibrationAddin.Host.Services;

namespace ExcelCalibrationAddin.Host.Generation
{
    /// <summary>
    /// Resolves the live workbook values for every mapped row.  Recognition
    /// supplies only ranges and relationships; this resolver is deliberately
    /// called after a fresh snapshot has been captured.
    /// </summary>
    public sealed class RowGenerationRuleResolver
    {
        private static readonly Regex NumberRegex = new Regex(
            @"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?",
            RegexOptions.Compiled);

        private readonly MeasurementRuleParameterResolver _parameterResolver;
        private readonly MeasurementRuleStructureAnalyzer _structureAnalyzer;
        private readonly NumberFormatInterpreter _formatInterpreter;

        public RowGenerationRuleResolver(
            MeasurementRuleParameterResolver parameterResolver = null,
            MeasurementRuleStructureAnalyzer structureAnalyzer = null,
            NumberFormatInterpreter formatInterpreter = null)
        {
            _parameterResolver = parameterResolver ?? new MeasurementRuleParameterResolver();
            _structureAnalyzer = structureAnalyzer ?? new MeasurementRuleStructureAnalyzer();
            _formatInterpreter = formatInterpreter ?? new NumberFormatInterpreter();
        }

        public IReadOnlyList<ResolvedRowRule> Resolve(WorkbookSnapshot snapshot, MeasurementRule rule)
        {
            var result = new List<ResolvedRowRule>();
            if (snapshot == null || rule == null)
            {
                return result;
            }

            var mappings = (rule.RowMappings ?? new List<MeasurementRowMapping>())
                .Where(mapping => mapping != null && mapping.Row > 0)
                .OrderBy(mapping => mapping.Row)
                .ToList();
            if (mappings.Count == 0)
            {
                mappings = (rule.WritableCells ?? new List<CellAddress>())
                    .Where(cell => cell != null && cell.Row > 0)
                    .GroupBy(cell => cell.Row)
                    .OrderBy(group => group.Key)
                    .Select((group, index) => new MeasurementRowMapping
                    {
                        Row = group.Key,
                        RowOrdinal = index + 1,
                        MeasurementCells = group.ToList(),
                        StandardValueRange = SelectRangeForRow(rule.StandardValueSource?.Range, group.Key),
                        ErrorRange = SelectRangeForRow(rule.ErrorSource?.Range, group.Key),
                        TechnicalRequirementRange = SelectRangeForRow(rule.MpeSource?.Range, group.Key),
                        AssociationKey = BuildAssociationKey(rule, group.Key, index + 1)
                    })
                    .ToList();
            }

            foreach (var mapping in mappings)
            {
                result.Add(ResolveRow(snapshot, rule, mapping));
            }

            return result;
        }

        private ResolvedRowRule ResolveRow(
            WorkbookSnapshot snapshot,
            MeasurementRule source,
            MeasurementRowMapping mapping)
        {
            var resolved = new ResolvedRowRule
            {
                Row = mapping.Row,
                MeasurementCells = (mapping.MeasurementCells ?? new List<CellAddress>())
                    .Where(cell => cell != null)
                    .Select(cell => new CellAddress { Row = cell.Row, Column = cell.Column })
                    .ToList(),
                AssociationKey = string.IsNullOrWhiteSpace(mapping.AssociationKey)
                    ? BuildAssociationKey(source, mapping.Row, mapping.RowOrdinal ?? 0)
                    : mapping.AssociationKey
            };

            var sheetName = source.TargetRange?.SheetName ?? mapping.StandardValueRange?.SheetName;
            var sheet = snapshot.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
            {
                return Fail(resolved, $"未读取到工作表“{sheetName}”。");
            }

            var rowRule = MeasurementRuleCloner.Clone(source);
            rowRule.TargetRange = RowRange(source.TargetRange, mapping.Row, mapping.MeasurementCells);
            rowRule.WritableCells = resolved.MeasurementCells.ToList();
            ApplyRange(rowRule.StandardValueSource, mapping.StandardValueRange);
            ApplyRange(rowRule.AverageSource, mapping.AverageRange);
            ApplyRange(rowRule.ErrorSource, mapping.ErrorRange);
            ApplyRange(rowRule.MpeSource, mapping.TechnicalRequirementRange);
            ApplyRange(rowRule.RangeSource, mapping.RangeValueRange);
            ApplyRange(rowRule.UncertaintySource, mapping.UncertaintyRange);
            ApplyRange(rowRule.ResultSource, mapping.ResultRange);
            rowRule.AdditionalJudgementConstraints = MeasurementRuleCloner.CloneJudgementConstraints(
                mapping.AdditionalJudgementConstraints ?? source.AdditionalJudgementConstraints);

            // Force fresh formula/requirement parsing for this row.  The
            // structure analyzer reads the current formula at the saved
            // address and recomputes dependencies/classification.
            _structureAnalyzer.Apply(snapshot, new[] { rowRule });
            rowRule = _parameterResolver.Apply(snapshot, new[] { rowRule }).FirstOrDefault() ?? rowRule;

            var targetCells = resolved.MeasurementCells
                .Select(cell => FindCell(sheet, cell.Row, cell.Column))
                .Where(cell => cell != null)
                .ToList();
            var targetFormatCell = targetCells.FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.NumberFormat))
                ?? targetCells.FirstOrDefault();
            var formatText = targetFormatCell?.NumberFormat ?? string.Empty;
            resolved.MeasurementFormat = _formatInterpreter.Interpret(formatText);
            if (!resolved.MeasurementFormat.DecimalPlaces.HasValue)
            {
                resolved.MeasurementFormat.DecimalPlaces = 1;
                resolved.PrecisionSource = "fallback:1";
            }
            else
            {
                resolved.PrecisionSource = "target.NumberFormat";
            }

            resolved.StandardValue = ResolveNumberInRange(sheet, mapping.StandardValueRange ?? rowRule.StandardValueSource?.Range);
            if (!resolved.StandardValue.HasValue && rowRule.FixedStandardValue.HasValue)
            {
                resolved.StandardValue = rowRule.FixedStandardValue.Value;
            }
            resolved.ErrorFormula = rowRule.ErrorFormula;
            resolved.ErrorType = GenerationRuleValidator.ResolveGenerationErrorType(rowRule);
            resolved.NegativeTolerance = rowRule.FixedNegativeTolerance ?? rowRule.FixedMpe;
            resolved.PositiveTolerance = rowRule.FixedPositiveTolerance ?? rowRule.FixedMpe;
            resolved.RequirementOperator = rowRule.RequirementOperator;
            resolved.ReferenceRange = rowRule.FixedReferenceRange;
            resolved.Unit = ResolveUnit(sheet, mapping, rowRule);
            resolved.RequiresExcelEvaluation = FormulaReferencesTarget(rowRule.ErrorFormula) ||
                FormulaReferencesTargetRequirement(rowRule.ErrorFormula, rowRule);

            if (resolved.MeasurementCells.Count == 0)
            {
                return Fail(resolved, "当前行没有可写测量值单元格。");
            }
            if (!resolved.StandardValue.HasValue)
            {
                return Fail(resolved, "当前标准值为空或无法解析。");
            }
            if (!resolved.NegativeTolerance.HasValue && !resolved.PositiveTolerance.HasValue)
            {
                return Fail(resolved, "当前技术要求为空或无法解析。");
            }

            resolved.IsValid = true;
            return resolved;
        }

        private static ResolvedRowRule Fail(ResolvedRowRule rule, string reason)
        {
            rule.IsValid = false;
            rule.FailureReason = reason;
            return rule;
        }

        private static bool FormulaReferencesTarget(ErrorFormulaInfo formula)
        {
            return formula != null && formula.ReferencesMeasurement;
        }

        private static bool FormulaReferencesTargetRequirement(ErrorFormulaInfo formula, MeasurementRule rule)
        {
            if (formula == null || string.IsNullOrWhiteSpace(formula.TechnicalRequirementFormula)) return false;
            var target = rule?.TargetRange;
            if (target == null) return false;
            var rowText = target.StartRow.ToString(CultureInfo.InvariantCulture);
            return formula.TechnicalRequirementFormula.IndexOf(rowText, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string ResolveUnit(SheetSnapshot sheet, MeasurementRowMapping mapping, MeasurementRule rule)
        {
            foreach (var range in new[]
            {
                mapping.TechnicalRequirementRange,
                mapping.ErrorRange,
                mapping.StandardValueRange,
                rule?.TargetRange
            }.Where(item => item != null))
            {
                foreach (var cell in CellsInRange(sheet, range))
                {
                    var unit = TemplateUnitParser.Extract(cell.DisplayText, cell.Text, cell.RawValueText, cell.NumberFormat);
                    if (!string.IsNullOrWhiteSpace(unit)) return unit;
                }
            }

            return string.Empty;
        }

        private static double? ResolveNumberInRange(SheetSnapshot sheet, CellRange range)
        {
            foreach (var cell in CellsInRange(sheet, range))
            {
                foreach (var text in new[] { cell.RawValueText, cell.Text, cell.DisplayText })
                {
                    var match = NumberRegex.Match(text ?? string.Empty);
                    if (!match.Success) continue;
                    if (double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) &&
                        !double.IsNaN(value) && !double.IsInfinity(value)) return value;
                }
            }
            return null;
        }

        private static IEnumerable<CellMeta> CellsInRange(SheetSnapshot sheet, CellRange range)
        {
            return (sheet?.Cells ?? new List<CellMeta>())
                .Where(cell => cell != null && range != null &&
                    cell.Row >= range.StartRow && cell.Row <= range.EndRow &&
                    cell.Column >= range.StartColumn && cell.Column <= range.EndColumn)
                .OrderBy(cell => cell.Row).ThenBy(cell => cell.Column);
        }

        private static CellMeta FindCell(SheetSnapshot sheet, int row, int column)
        {
            return sheet?.Cells?.FirstOrDefault(cell => cell.Row == row && cell.Column == column);
        }

        private static void ApplyRange(ParameterSource source, CellRange range)
        {
            if (source != null && range != null) source.Range = CloneRange(range);
        }

        private static CellRange RowRange(CellRange source, int row, IReadOnlyList<CellAddress> cells)
        {
            var start = cells?.Count > 0 ? cells.Min(cell => cell.Column) : source?.StartColumn ?? 1;
            var end = cells?.Count > 0 ? cells.Max(cell => cell.Column) : source?.EndColumn ?? start;
            return new CellRange
            {
                SheetName = source?.SheetName ?? string.Empty,
                StartRow = row, EndRow = row, StartColumn = start, EndColumn = end
            };
        }

        private static CellRange SelectRangeForRow(CellRange source, int row)
        {
            if (source == null || row < source.StartRow || row > source.EndRow) return source;
            return new CellRange
            {
                SheetName = source.SheetName, StartRow = row, EndRow = row,
                StartColumn = source.StartColumn, EndColumn = source.EndColumn
            };
        }

        private static string BuildAssociationKey(MeasurementRule rule, int row, int ordinal)
        {
            return string.Join("|", new[]
            {
                rule?.FieldName ?? string.Empty,
                rule?.BlockOrdinal?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                ordinal.ToString(CultureInfo.InvariantCulture),
                row.ToString(CultureInfo.InvariantCulture)
            });
        }

        private static CellRange CloneRange(CellRange range)
        {
            return range == null ? null : new CellRange
            {
                SheetName = range.SheetName, StartRow = range.StartRow, EndRow = range.EndRow,
                StartColumn = range.StartColumn, EndColumn = range.EndColumn
            };
        }
    }
}
