using System;
using System.Diagnostics;
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
    public sealed class FormulaResultVerifier
    {
        private static readonly Regex NumberRegex = new Regex(@"[-+]?\d+(\.\d+)?([eE][-+]?\d+)?", RegexOptions.Compiled);

        public void Verify(WorkbookSnapshot snapshot, IReadOnlyList<MeasurementRule> rules)
        {
            if (snapshot == null || rules == null)
            {
                return;
            }

            new MeasurementRuleStructureAnalyzer().Apply(snapshot, rules);
            foreach (var rule in rules.Where(item => item != null &&
                                                    (item.ErrorFormula?.HasFormula == true ||
                                                    GenerationRuleValidator.IsRepeatabilityRule(item))))
            {
                GenerationRuleValidator.ValidateFormulaDependencies(rule);
                VerifyRepeatabilityValues(snapshot, rule);
                if (rule.ErrorFormula?.HasFormula == true)
                {
                    VerifyRule(snapshot, rule);
                }

                foreach (var constraint in rule.AdditionalJudgementConstraints ?? new List<MeasurementJudgementConstraint>())
                {
                    if (constraint?.ErrorFormula?.HasFormula != true)
                    {
                        continue;
                    }

                    var overlay = MeasurementRuleCloner.Clone(rule);
                    overlay.AdditionalJudgementConstraints = new List<MeasurementJudgementConstraint>();
                    MeasurementJudgementConstraintHelper.Overlay(overlay, constraint);
                    VerifyRule(snapshot, overlay);
                }
            }
        }

        private static void VerifyRule(WorkbookSnapshot snapshot, MeasurementRule rule)
        {
            if (!GenerationRuleValidator.HasValidRange(rule.ErrorSource?.Range))
            {
                return;
            }

            var sheet = snapshot.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, rule.ErrorSource.Range.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
            {
                throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”公式验证失败：未读取到误差公式所在工作表。");
            }

            var formulaCells = MergedCellLogicalRangeResolver.GetContentCells(sheet, rule.ErrorSource.Range)
                .Select(item => item.Anchor)
                .Where(item => item != null)
                .ToList();
            // ErrorSource often spans the whole visual block (including merged or
            // blank rows), while generation only writes rows that have a writable
            // measurement cell.  Verifying every formula in the broad range can
            // therefore validate stale formulas from non-generated rows.  Prefer
            // the row mappings (or writable cells) as the authoritative set of
            // generated rows when that information is available.
            var generatedRows = ResolveGeneratedRows(rule);
            if (generatedRows.Count > 0)
            {
                formulaCells = formulaCells
                    .Where(cell => IsFormulaForGeneratedRow(cell, generatedRows))
                    .ToList();
            }
            if (formulaCells.Count == 0)
            {
                throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”公式验证失败：误差区域没有可读取的公式结果。");
            }

            var bounds = ResolveFormulaResultBounds(rule);
            var requirementOperator = GenerationRuleValidator.IsUpperLimitRule(rule)
                ? GenerationRuleValidator.ResolveUpperLimitOperator(rule)
                : rule.RequirementOperator;
            if (GenerationRuleValidator.IsUpperLimitRule(rule))
            {
                bounds = (
                    requirementOperator == TechnicalRequirementOperator.GreaterThan ||
                    requirementOperator == TechnicalRequirementOperator.GreaterThanOrEqual
                        ? Math.Abs(rule.FixedMpe.GetValueOrDefault())
                        : 0,
                    Math.Abs(rule.FixedMpe.GetValueOrDefault()));
            }
            foreach (var cell in formulaCells)
            {
                if (string.IsNullOrWhiteSpace(cell.Formula))
                {
                    throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”公式验证失败：R{cell.Row}C{cell.Column} 不再是公式单元格。");
                }

                double value;
                if (!TryReadFormulaResult(cell, rule.ErrorFormula, out value))
                {
                    throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”公式验证失败：R{cell.Row}C{cell.Column} 的公式结果不可读。");
                }

                if (IsDisplayedAsZero(value, cell))
                {
                    if (GenerationRuleValidator.IsRepeatabilityRule(rule))
                    {
                        throw new RepeatabilityVerificationException(
                            $"“{GenerationRuleValidator.ResolveRuleName(rule)}”生成后重复性在误差列分辨力下显示为 0。\n" +
                            "请提高测量值小数位数或放宽重复性生成区间。");
                    }

                    throw new DisplayedErrorZeroVerificationException(
                        $"“{GenerationRuleValidator.ResolveRuleName(rule)}”生成后误差在 R{cell.Row}C{cell.Column} 的当前分辨力下显示为 0。");
                }

                if (!IsRequirementSatisfied(requirementOperator, value, bounds) &&
                    (requirementOperator != TechnicalRequirementOperator.None ||
                     value < bounds.lower - 1e-12 || value > bounds.upper + 1e-12))
                {
                    var diagnostic = BuildOutOfRangeDiagnostic(snapshot, rule, cell, value, requirementOperator, bounds);
                    AddinFileLogger.Configure("VSTO");
                    Trace.WriteLine("[Generation][FormulaOutOfRange] " + diagnostic);
                    throw new InvalidOperationException(
                        $"{diagnostic}{Environment.NewLine}{Environment.NewLine}详细错误已记录到：{AddinFileLogger.LogFilePath}{Environment.NewLine}请复制以上“公式超限诊断”内容反馈排查。");
                }
            }
        }

        private static string BuildOutOfRangeDiagnostic(WorkbookSnapshot snapshot, MeasurementRule rule, CellMeta cell, double value, TechnicalRequirementOperator requirementOperator, (double lower, double upper) bounds)
        {
            var ruleName = GenerationRuleValidator.ResolveRuleName(rule);
            var range = rule.ErrorSource?.Range;
            var workbook = string.IsNullOrWhiteSpace(snapshot?.WorkbookName) ? "未知" : snapshot.WorkbookName;
            return string.Join(Environment.NewLine, new[]
            {
                "公式超限诊断:",
                $"项目={ruleName}",
                $"工作簿={workbook}",
                $"工作表={range?.SheetName ?? "未知"}",
                $"公式单元格=R{cell?.Row ?? 0}C{cell?.Column ?? 0}",
                $"公式={cell?.Formula ?? ""}",
                $"公式结果={FormatDiagnosticNumber(value)}",
                $"原始结果={cell?.RawValueText ?? ""}",
                $"显示结果={cell?.Text ?? cell?.DisplayText ?? ""}",
                $"技术要求操作符={requirementOperator}",
                $"允许范围=[{FormatDiagnosticNumber(bounds.lower)}, {FormatDiagnosticNumber(bounds.upper)}]",
                $"误差类型={rule.ErrorType}",
                $"固定技术要求={FormatDiagnosticNumber(rule.FixedMpe)}",
                $"负向限值={FormatDiagnosticNumber(rule.FixedNegativeTolerance)}",
                $"正向限值={FormatDiagnosticNumber(rule.FixedPositiveTolerance)}",
                $"测量值区域={rule.TargetRange?.ToString() ?? "未知"}",
                $"误差区域={range?.ToString() ?? "未知"}"
            });
        }

        private static string FormatDiagnosticNumber(double? value)
        {
            return value.HasValue ? value.Value.ToString("G17", CultureInfo.InvariantCulture) : "无";
        }

        private static void VerifyRepeatabilityValues(WorkbookSnapshot snapshot, MeasurementRule rule)
        {
            if (!GenerationRuleValidator.IsRepeatabilityRule(rule) ||
                !GenerationRuleValidator.HasValidRange(rule.TargetRange))
            {
                return;
            }

            var sheet = snapshot.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, rule.TargetRange.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
            {
                throw new InvalidOperationException(
                    $"“{GenerationRuleValidator.ResolveRuleName(rule)}”生成后验证失败：未读取到测量值工作表。");
            }

            var values = MergedCellLogicalRangeResolver.GetContentCells(sheet, rule.TargetRange)
                .Select(item => item.Anchor)
                .Where(item => item != null)
                .Select(cell =>
                {
                    double value;
                    return TryReadRawFormulaResult(cell.RawValueText, out value) ||
                           TryReadRawFormulaResult(cell.Text, out value)
                        ? (double?)value
                        : null;
                })
                .Where(value => value.HasValue)
                .Select(value => value.Value)
                .ToList();
            if (values.Count < 2 || values.Distinct().Count() < 2)
            {
                throw new RepeatabilityVerificationException(
                    $"“{GenerationRuleValidator.ResolveRuleName(rule)}”生成后重复性为 0：写入后的测量值没有可见波动。");
            }
        }

        private static bool IsDisplayedAsZero(double value, CellMeta cell)
        {
            var decimalPlaces = new NumberFormatInterpreter().Interpret(cell?.NumberFormat).DecimalPlaces;
            if (!decimalPlaces.HasValue)
            {
                return Math.Abs(value) <= 1e-12;
            }

            return Math.Abs(Math.Round(value, Math.Max(0, Math.Min(15, decimalPlaces.Value)))) <= 1e-12;
        }

        private static bool IsRequirementSatisfied(
            TechnicalRequirementOperator requirementOperator,
            double value,
            (double lower, double upper) bounds)
        {
            var magnitude = Math.Abs(value);
            var limit = Math.Max(Math.Abs(bounds.lower), Math.Abs(bounds.upper));
            var tolerance = ComparisonTolerance(limit);
            switch (requirementOperator)
            {
                case TechnicalRequirementOperator.LessThan:
                    return magnitude < limit || (magnitude > limit && magnitude - limit <= tolerance);
                case TechnicalRequirementOperator.LessThanOrEqual:
                case TechnicalRequirementOperator.PlusMinus:
                    return magnitude <= limit + tolerance;
                case TechnicalRequirementOperator.GreaterThan:
                    return magnitude > limit || (magnitude < limit && limit - magnitude <= tolerance);
                case TechnicalRequirementOperator.GreaterThanOrEqual:
                    return magnitude >= limit - tolerance;
                default:
                    return false;
            }
        }

        private static double ComparisonTolerance(double limit)
        {
            return Math.Max(1e-12, Math.Abs(limit) * 1e-12);
        }

        private static (double lower, double upper) ResolveFormulaResultBounds(MeasurementRule rule)
        {
            var negativeTolerance = Math.Abs(rule.FixedNegativeTolerance ?? rule.FixedMpe.GetValueOrDefault());
            var positiveTolerance = Math.Abs(rule.FixedPositiveTolerance ?? rule.FixedMpe.GetValueOrDefault());
            switch (rule.ErrorFormula?.Scale)
            {
                case ErrorFormulaScale.RelativeToStandardValue:
                case ErrorFormulaScale.RelativeToReferenceRange:
                    return (
                        -Math.Abs(ScaleRatio(negativeTolerance, rule.ErrorFormula)),
                        Math.Abs(ScaleRatio(positiveTolerance, rule.ErrorFormula)));
                default:
                    switch (rule.ErrorType)
                    {
                        case ErrorType.Relative:
                            return (
                                -Math.Abs(rule.FixedStandardValue.GetValueOrDefault() * negativeTolerance),
                                Math.Abs(rule.FixedStandardValue.GetValueOrDefault() * positiveTolerance));
                        case ErrorType.Referenced:
                            return (
                                -Math.Abs(rule.FixedReferenceRange.GetValueOrDefault() * negativeTolerance),
                                Math.Abs(rule.FixedReferenceRange.GetValueOrDefault() * positiveTolerance));
                        default:
                            return (-negativeTolerance, positiveTolerance);
                    }
            }
        }

        private static double ScaleRatio(double ratio, ErrorFormulaInfo formula)
        {
            return formula?.FormulaMultipliesBy100 == true ? ratio * 100.0 : ratio;
        }

        private static bool TryReadFormulaResult(CellMeta cell, ErrorFormulaInfo formula, out double value)
        {
            value = 0;
            if (TryReadRawFormulaResult(cell?.RawValueText, out value))
            {
                return true;
            }

            var text = cell?.Text;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var match = NumberRegex.Match(text.Replace(",", string.Empty));
            if (!match.Success ||
                !double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
            {
                return false;
            }

            if ((formula?.Scale == ErrorFormulaScale.RelativeToStandardValue ||
                 formula?.Scale == ErrorFormulaScale.RelativeToReferenceRange) &&
                text.Contains("%") &&
                formula?.FormulaMultipliesBy100 != true)
            {
                value /= 100.0;
            }

            return true;
        }

        private static bool TryReadRawFormulaResult(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return double.TryParse(
                       text,
                       NumberStyles.Float | NumberStyles.AllowThousands,
                       CultureInfo.InvariantCulture,
                       out value) ||
                   double.TryParse(
                       text,
                       NumberStyles.Float | NumberStyles.AllowThousands,
                       CultureInfo.CurrentCulture,
                       out value);
        }

        private static HashSet<int> ResolveGeneratedRows(MeasurementRule rule)
        {
            var rows = new HashSet<int>();
            foreach (var mapping in rule?.RowMappings ?? new List<MeasurementRowMapping>())
            {
                if (mapping != null && mapping.Row > 0 &&
                    (mapping.MeasurementCells?.Count ?? 0) > 0)
                {
                    rows.Add(mapping.Row);
                }
            }

            if (rows.Count == 0)
            {
                foreach (var cell in rule?.WritableCells ?? new List<CellAddress>())
                {
                    if (cell?.Row > 0)
                    {
                        rows.Add(cell.Row);
                    }
                }
            }

            return rows;
        }

        private static bool IsFormulaForGeneratedRow(CellMeta cell, ISet<int> generatedRows)
        {
            if (cell == null || generatedRows == null || generatedRows.Count == 0)
            {
                return true;
            }

            if (generatedRows.Contains(cell.Row))
            {
                return true;
            }

            // A merged formula cell can have an anchor row different from the
            // row displayed in the logical block. Match the rows referenced by
            // the formula as a fallback.
            foreach (Match match in Regex.Matches(cell.Formula ?? string.Empty, @"\$?[A-Z]{1,3}\$?(?<row>\d+)", RegexOptions.IgnoreCase))
            {
                if (int.TryParse(match.Groups["row"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row) &&
                    generatedRows.Contains(row))
                {
                    return true;
                }
            }

            return false;
        }
    }

    public class GenerationRetryException : InvalidOperationException
    {
        public GenerationRetryException(string message)
            : base(message)
        {
        }
    }

    public sealed class RepeatabilityVerificationException : GenerationRetryException
    {
        public RepeatabilityVerificationException(string message)
            : base(message)
        {
        }
    }

    public sealed class DisplayedErrorZeroVerificationException : GenerationRetryException
    {
        public DisplayedErrorZeroVerificationException(string message)
            : base(message)
        {
        }
    }
}
