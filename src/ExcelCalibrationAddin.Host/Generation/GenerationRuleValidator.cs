using System;
using ExcelCalibrationAddin.Contracts;

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

        private static bool ContainsRuleName(MeasurementRule rule, string keyword)
        {
            if (rule == null || string.IsNullOrWhiteSpace(keyword))
            {
                return false;
            }

            return (rule.FieldName ?? string.Empty).IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                (rule.FieldAlias ?? string.Empty).IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsUpperLimitRule(MeasurementRule rule)
        {
            return ResolveRuleName(rule).IndexOf("响应时间", StringComparison.OrdinalIgnoreCase) >= 0;
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
            if (rule?.TargetRange == null)
            {
                throw new InvalidOperationException($"“{ResolveRuleName(rule)}”未设置测量值写入区域。");
            }

            if (writableCellCount <= 0)
            {
                throw new InvalidOperationException(AppendReason($"“{ResolveRuleName(rule)}”的测量值写入区域无效。", writableFailureReason));
            }
        }

        private static string AppendReason(string message, string reason)
        {
            return string.IsNullOrWhiteSpace(reason)
                ? message
                : $"{message}{Environment.NewLine}原因：{reason}";
        }
    }
}
