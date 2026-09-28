using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Services;

namespace ExcelCalibrationAddin.Host.Generation
{
    public static class GenerationRuleValidator
    {
        public static string ResolveRuleName(MeasurementRule rule)
        {
            var ruleName = string.IsNullOrWhiteSpace(rule?.FieldAlias)
                ? rule?.FieldName
                : rule.FieldAlias;
            return string.IsNullOrWhiteSpace(ruleName) ? "未命名校准项" : ruleName;
        }

        public static ErrorType ResolveGenerationErrorType(MeasurementRule rule)
        {
            switch (rule?.ErrorFormula?.Scale)
            {
                case ErrorFormulaScale.RelativeToReferenceRange:
                    return ErrorType.Referenced;
                case ErrorFormulaScale.RelativeToStandardValue:
                    return ErrorType.Relative;
                default:
                    return rule?.ErrorType ?? ErrorType.Absolute;
            }
        }

        public static bool HasValidRange(CellRange range)
        {
            return range != null &&
                !string.IsNullOrWhiteSpace(range.SheetName) &&
                range.StartRow > 0 &&
                range.StartColumn > 0 &&
                range.EndRow >= range.StartRow &&
                range.EndColumn >= range.StartColumn;
        }

        public static bool IsRepeatabilityRule(MeasurementRule rule)
        {
            return ContainsRuleName(rule, "重复性");
        }

        public static bool IsRepeatabilityGenerationRule(MeasurementRule rule)
        {
            if (!IsRepeatabilityRule(rule))
            {
                return false;
            }

            var formula = rule.ErrorFormula?.Formula;
            if (string.IsNullOrWhiteSpace(formula))
            {
                return true;
            }

            // A combined item can mention repeatability in its title while its
            // primary formula calculates a different quantity (for example,
            // indication error). Use the dedicated repeatability generator only
            // when the primary formula itself computes a spread/dispersion metric.
            return Regex.IsMatch(formula, @"STDEV|VAR\s*\(|MAX\s*\([^)]*\)\s*-\s*MIN\s*\(",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static bool ContainsRuleName(MeasurementRule rule, string keyword)
        {
            if (rule == null || string.IsNullOrWhiteSpace(keyword))
            {
                return false;
            }

            return (rule.FieldName ?? string.Empty).IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                (rule.FieldAlias ?? string.Empty).IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsTimeNamedRule(MeasurementRule rule)
        {
            return ContainsRuleName(rule, "时间");
        }

        public static bool IsUpperLimitRule(MeasurementRule rule)
        {
            if (!IsTimeNamedRule(rule))
            {
                return false;
            }

            var requirementOperator = ResolveUpperLimitOperator(rule);
            return ResolveUpperLimitUnit(rule) == "s" &&
                (requirementOperator == TechnicalRequirementOperator.LessThan ||
                 requirementOperator == TechnicalRequirementOperator.LessThanOrEqual ||
                 requirementOperator == TechnicalRequirementOperator.GreaterThan ||
                 requirementOperator == TechnicalRequirementOperator.GreaterThanOrEqual);
        }

        public static TechnicalRequirementOperator ResolveUpperLimitOperator(MeasurementRule rule)
        {
            var requirementOperator = rule?.RequirementOperator ?? TechnicalRequirementOperator.None;
            var pattern = MpeValuePatternCodec.Parse(rule?.MpeSource?.ValuePattern);
            if (requirementOperator == TechnicalRequirementOperator.None && pattern != null)
            {
                requirementOperator = pattern.Operator;
            }

            if (requirementOperator == TechnicalRequirementOperator.None)
            {
                requirementOperator = InferOperatorFromResultFormula(rule);
            }

            return requirementOperator;
        }

        public static string ResolveUpperLimitUnit(MeasurementRule rule)
        {
            var pattern = MpeValuePatternCodec.Parse(rule?.MpeSource?.ValuePattern);
            var unit = pattern?.Unit;
            if (string.IsNullOrWhiteSpace(unit))
            {
                unit = MpeValuePatternCodec.NormalizeUnit(rule?.FormatRule?.UnitSuffix);
            }

            if (string.IsNullOrWhiteSpace(unit))
            {
                unit = MpeValuePatternCodec.NormalizeUnit(ResolveTemplateRequirementUnit(rule));
            }

            return unit ?? string.Empty;
        }

        public static TechnicalRequirementOperator InferOperatorFromResultFormula(MeasurementRule rule)
        {
            var formula = rule?.ErrorFormula?.ResultFormula;
            if (string.IsNullOrWhiteSpace(formula))
            {
                return TechnicalRequirementOperator.None;
            }

            var fallbackSheetName = rule.ResultSource?.Range?.SheetName ??
                rule.MpeSource?.Range?.SheetName ??
                rule.ErrorSource?.Range?.SheetName ??
                rule.TargetRange?.SheetName ??
                string.Empty;
            foreach (Match match in ResultComparisonRegex.Matches(formula))
            {
                var requirementOperator = ParseComparisonOperator(match.Groups["op"].Value);
                if (requirementOperator == TechnicalRequirementOperator.None)
                {
                    continue;
                }

                var leftRanges = TemplateFormulaParser.ExtractReferencedRanges(
                    match.Groups["left"].Value,
                    fallbackSheetName);
                var rightRanges = TemplateFormulaParser.ExtractReferencedRanges(
                    match.Groups["right"].Value,
                    fallbackSheetName);
                var leftIsRequirement = leftRanges.Any(range =>
                    RangesOverlap(range, rule.MpeSource?.Range));
                var rightIsRequirement = rightRanges.Any(range =>
                    RangesOverlap(range, rule.MpeSource?.Range));
                var leftIsMeasurement = leftRanges.Any(range =>
                    RangesOverlap(range, rule.ErrorSource?.Range) ||
                    RangesOverlap(range, rule.AverageSource?.Range) ||
                    RangesOverlap(range, rule.TargetRange));
                var rightIsMeasurement = rightRanges.Any(range =>
                    RangesOverlap(range, rule.ErrorSource?.Range) ||
                    RangesOverlap(range, rule.AverageSource?.Range) ||
                    RangesOverlap(range, rule.TargetRange));
                if (leftIsMeasurement && rightIsRequirement)
                {
                    return requirementOperator;
                }

                if (leftIsRequirement && rightIsMeasurement)
                {
                    return InvertComparisonOperator(requirementOperator);
                }
            }

            return TechnicalRequirementOperator.None;
        }

        public static bool IsNonNumericRule(MeasurementRule rule)
        {
            return ResolveRuleName(rule).IndexOf("外观", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsAlarmRule(MeasurementRule rule)
        {
            return ResolveRuleName(rule).IndexOf("报警", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static void ValidateAlarmRule(MeasurementRule rule, int writableCellCount, string writableFailureReason = null)
        {
            ValidateCommonWritableRule(rule, writableCellCount, writableFailureReason);
            if (!rule.FixedStandardValue.HasValue)
            {
                throw new InvalidOperationException("请先在功能区的“报警值输入”中输入具体数值。");
            }
        }

        public static void ValidateRule(MeasurementRule rule, int writableCellCount, string writableFailureReason = null)
        {
            ValidateFormulaDependencies(rule);
            var ruleName = ResolveRuleName(rule);
            if (rule.TargetRange == null)
            {
                throw new InvalidOperationException($"“{ruleName}”未设置测量值写入区域。");
            }

            if (!rule.FixedMpe.HasValue || rule.FixedMpe.Value <= 0)
            {
                throw new InvalidOperationException($"“{ruleName}”缺少有效的允许误差。请检查技术要求区域或模板规则。");
            }

            if (!HasValidRange(rule.ErrorSource?.Range))
            {
                throw new InvalidOperationException($"“{ruleName}”缺少误差区域。请在侧边栏设置误差区域，或删除无需生成的校准项。");
            }

            if (!rule.FixedStandardValue.HasValue && !HasValidRange(rule.StandardValueSource?.Range))
            {
                throw new InvalidOperationException($"“{ruleName}”缺少标准值。请检查标准值区域，或在侧边栏设置手动标准值。");
            }

            if (ResolveGenerationErrorType(rule) == ErrorType.Referenced &&
                (!rule.FixedReferenceRange.HasValue || rule.FixedReferenceRange.Value <= 0))
            {
                throw new InvalidOperationException($"“{ruleName}”使用引用误差时必须提供有效量程。");
            }

            if (writableCellCount <= 0)
            {
                throw new InvalidOperationException(AppendReason($"“{ruleName}”的测量值写入区域无效。", writableFailureReason));
            }

            rule.GroupSize = writableCellCount;
        }

        public static void ValidateRepeatabilityRule(MeasurementRule rule, int writableCellCount, string writableFailureReason = null)
        {
            ValidateCommonWritableRule(rule, writableCellCount, writableFailureReason);
            if (writableCellCount < 2)
            {
                throw new InvalidOperationException(
                    "Repeatability generation requires at least two writable measurement cells to produce a non-zero result.");
            }

            if (!rule.FixedStandardValue.HasValue)
            {
                throw new InvalidOperationException($"“{ResolveRuleName(rule)}”缺少标准值，无法生成重复性测量值。");
            }

            if (!rule.FixedMpe.HasValue || rule.FixedMpe.Value <= 0)
            {
                throw new InvalidOperationException($"“{ResolveRuleName(rule)}”缺少有效的重复性技术要求。");
            }
        }

        public static void ValidateUpperLimitRule(MeasurementRule rule, int writableCellCount, string writableFailureReason = null)
        {
            ValidateCommonWritableRule(rule, writableCellCount, writableFailureReason);
            if (!rule.FixedMpe.HasValue || rule.FixedMpe.Value <= 0)
            {
                throw new InvalidOperationException($"“{ResolveRuleName(rule)}”缺少有效的上限技术要求。");
            }
        }

        private static void ValidateCommonWritableRule(MeasurementRule rule, int writableCellCount, string writableFailureReason)
        {
            ValidateFormulaDependencies(rule);
            if (rule?.TargetRange == null)
            {
                throw new InvalidOperationException($"“{ResolveRuleName(rule)}”未设置测量值写入区域。");
            }

            if (writableCellCount <= 0)
            {
                throw new InvalidOperationException(AppendReason($"“{ResolveRuleName(rule)}”的测量值写入区域无效。", writableFailureReason));
            }
        }

        public static void ValidateFormulaDependencies(MeasurementRule rule)
        {
            if (!string.IsNullOrWhiteSpace(rule?.RecognitionError))
            {
                throw new InvalidOperationException(
                    $"“{ResolveRuleName(rule)}”模板识别失败：{rule.RecognitionError}。请修正误差/MPE布局后重新识别。");
            }

            var unresolved = rule?.ErrorFormula?.UnresolvedDependencies?
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new System.Collections.Generic.List<string>();
            if (unresolved.Count > 0)
            {
                throw new InvalidOperationException(
                    $"“{ResolveRuleName(rule)}”包含当前无法安全解析的公式依赖：{string.Join("；", unresolved)}。" +
                    "请改为当前工作簿内的普通单元格引用后重新识别模板。");
            }

            foreach (var constraint in rule?.AdditionalJudgementConstraints ?? new System.Collections.Generic.List<MeasurementJudgementConstraint>())
            {
                var constraintUnresolved = constraint?.ErrorFormula?.UnresolvedDependencies?
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList() ?? new System.Collections.Generic.List<string>();
                if (constraintUnresolved.Count > 0)
                {
                    throw new InvalidOperationException(
                        $"“{ResolveRuleName(rule)}”的附属判定包含当前无法安全解析的公式依赖：{string.Join("；", constraintUnresolved)}。" +
                        "请改为当前工作簿内的普通单元格引用后重新识别模板。");
                }
            }
        }

        private static string AppendReason(string message, string reason)
        {
            return string.IsNullOrWhiteSpace(reason)
                ? message
                : $"{message}{Environment.NewLine}原因：{reason}";
        }

        private static readonly Regex ResultComparisonRegex = new Regex(
            @"(?<left>(?:(?:'[^']+'|[A-Za-z0-9_一-鿿]+)!)?\$?[A-Z]{1,3}\$?\d+(?::(?:(?:'[^']+'|[A-Za-z0-9_一-鿿]+)!)?\$?[A-Z]{1,3}\$?\d+)?)\s*(?<op><=|>=|<|>)\s*(?<right>(?:(?:'[^']+'|[A-Za-z0-9_一-鿿]+)!)?\$?[A-Z]{1,3}\$?\d+(?::(?:(?:'[^']+'|[A-Za-z0-9_一-鿿]+)!)?\$?[A-Z]{1,3}\$?\d+)?)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static TechnicalRequirementOperator ParseComparisonOperator(string text)
        {
            switch ((text ?? string.Empty).Trim())
            {
                case "<":
                    return TechnicalRequirementOperator.LessThan;
                case "<=":
                    return TechnicalRequirementOperator.LessThanOrEqual;
                case ">":
                    return TechnicalRequirementOperator.GreaterThan;
                case ">=":
                    return TechnicalRequirementOperator.GreaterThanOrEqual;
                default:
                    return TechnicalRequirementOperator.None;
            }
        }

        private static TechnicalRequirementOperator InvertComparisonOperator(TechnicalRequirementOperator requirementOperator)
        {
            switch (requirementOperator)
            {
                case TechnicalRequirementOperator.LessThan:
                    return TechnicalRequirementOperator.GreaterThan;
                case TechnicalRequirementOperator.LessThanOrEqual:
                    return TechnicalRequirementOperator.GreaterThanOrEqual;
                case TechnicalRequirementOperator.GreaterThan:
                    return TechnicalRequirementOperator.LessThan;
                case TechnicalRequirementOperator.GreaterThanOrEqual:
                    return TechnicalRequirementOperator.LessThanOrEqual;
                default:
                    return TechnicalRequirementOperator.None;
            }
        }

        private static bool RangesOverlap(CellRange left, CellRange right)
        {
            if (!HasValidRange(left) || !HasValidRange(right))
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(left.SheetName) &&
                !string.IsNullOrWhiteSpace(right.SheetName) &&
                !string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return left.StartRow <= right.EndRow &&
                right.StartRow <= left.EndRow &&
                left.StartColumn <= right.EndColumn &&
                right.StartColumn <= left.EndColumn;
        }

        private static string ResolveTemplateRequirementUnit(MeasurementRule rule)
        {
            var regions = rule?.TemplateDefinition?.Regions ?? new List<TemplateRegionDefinition>();
            foreach (var region in regions.Where(item =>
                item != null &&
                (item.Role == TemplateRegionRole.TechnicalRequirement ||
                 item.Role == TemplateRegionRole.MeasurementValue ||
                 item.Role == TemplateRegionRole.AverageValue)))
            {
                var unit = MpeValuePatternCodec.NormalizeUnit(region.Unit) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(unit))
                {
                    return unit;
                }

                foreach (var candidate in region.Units ?? new List<string>())
                {
                    unit = MpeValuePatternCodec.NormalizeUnit(candidate);
                    if (!string.IsNullOrWhiteSpace(unit))
                    {
                        return unit;
                    }
                }

                foreach (var requirement in region.RequirementValues ?? new List<TemplateRequirementValue>())
                {
                    unit = MpeValuePatternCodec.NormalizeUnit(requirement?.Unit);
                    if (!string.IsNullOrWhiteSpace(unit))
                    {
                        return unit;
                    }
                }
            }

            return string.Empty;
        }
    }
}
