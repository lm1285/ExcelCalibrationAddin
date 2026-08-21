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
                VerifyRepeatabilityValues(snapshot, rule);
                if (rule.ErrorFormula?.HasFormula == true)
                {
                    VerifyRule(snapshot, rule);
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
            if (formulaCells.Count == 0)
            {
                throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”公式验证失败：误差区域没有可读取的公式结果。");
            }

            var bounds = ResolveFormulaResultBounds(rule);
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

                if (!IsRequirementSatisfied(rule.RequirementOperator, value, bounds) &&
                    (rule.RequirementOperator != TechnicalRequirementOperator.None ||
                     value < bounds.lower - 1e-12 || value > bounds.upper + 1e-12))
                {
                    throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”生成后 Excel 公式结果超出技术要求。");
                }
            }
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
