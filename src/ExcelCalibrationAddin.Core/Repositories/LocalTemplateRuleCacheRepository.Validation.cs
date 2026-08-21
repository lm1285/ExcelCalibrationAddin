using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Repositories
{
    public sealed partial class LocalTemplateRuleCacheRepository
    {
        public void ValidateTemplateForSave(IReadOnlyList<MeasurementRule> rules)
        {
            ValidateSavableTemplate(rules);
        }

        private static void ValidateSavableTemplate(IReadOnlyList<MeasurementRule> rules)
        {
            if (rules == null || rules.Count == 0)
            {
                throw new InvalidOperationException("模板没有可保存的校准项规则。");
            }

            var missingMessages = new List<string>();
            foreach (var rule in rules.Where(item => item != null && item.IsEnabled))
            {
                if (IsNonNumericCalibrationItem(rule))
                {
                    continue;
                }

                var missing = new List<string>();
                if (!HasValidRange(rule.TargetRange))
                {
                    missing.Add("测量值");
                }

                if (IsAlarmCalibrationItem(rule))
                {
                    if (missing.Count > 0)
                    {
                        missingMessages.Add($"“{ResolveRuleName(rule)}”缺少{string.Join("、", missing)}");
                    }

                    continue;
                }

                if (!rule.FixedStandardValue.HasValue && !HasValidRange(rule.StandardValueSource?.Range))
                {
                    missing.Add("标准值");
                }

                if (!HasValidRange(rule.ErrorSource?.Range))
                {
                    missing.Add("误差");
                }

                if ((!rule.FixedMpe.HasValue || rule.FixedMpe.Value <= 0) && !HasValidRange(rule.MpeSource?.Range))
                {
                    missing.Add("技术要求/允许误差");
                }

                if (rule.TargetRange != null &&
                    rule.TargetRange.EndRow > rule.TargetRange.StartRow &&
                    !ErrorFormulaClassifier.IsMaximumError(rule) &&
                    (rule.RowMappings == null || rule.RowMappings.Count == 0 || rule.RowMappings.Any(item => item == null || !item.IsComplete)))
                {
                    missing.Add("行级映射");
                }

                if (missing.Count > 0)
                {
                    missingMessages.Add($"“{ResolveRuleName(rule)}”缺少{string.Join("、", missing)}");
                }
            }

            if (missingMessages.Count > 0)
            {
                throw new InvalidOperationException("模板必填字段不完整，无法保存：" + Environment.NewLine + string.Join(Environment.NewLine, missingMessages));
            }
        }

        private static bool HasValidRange(CellRange range)
        {
            return range != null &&
                !string.IsNullOrWhiteSpace(range.SheetName) &&
                range.StartRow > 0 &&
                range.StartColumn > 0 &&
                range.EndRow >= range.StartRow &&
                range.EndColumn >= range.StartColumn;
        }

        private static string ResolveRuleName(MeasurementRule rule)
        {
            var name = string.IsNullOrWhiteSpace(rule?.FieldAlias) ? rule?.FieldName : rule.FieldAlias;
            return string.IsNullOrWhiteSpace(name) ? "未命名校准项" : name;
        }

        private static bool IsNonNumericCalibrationItem(MeasurementRule rule)
        {
            var name = ResolveRuleName(rule);
            return name.IndexOf("外观", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsAlarmCalibrationItem(MeasurementRule rule)
        {
            var name = ResolveRuleName(rule);
            return name.IndexOf("报警", StringComparison.OrdinalIgnoreCase) >= 0;
        }

    }
}
