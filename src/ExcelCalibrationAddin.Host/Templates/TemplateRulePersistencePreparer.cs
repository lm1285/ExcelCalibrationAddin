using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Services;

namespace ExcelCalibrationAddin.Host.Templates
{
    internal static class TemplateRulePersistencePreparer
    {
        public static IReadOnlyList<MeasurementRule> Prepare(IEnumerable<MeasurementRule> rules)
        {
            return (rules ?? Enumerable.Empty<MeasurementRule>())
                .Where(rule => rule != null)
                .Select(MeasurementRuleCloner.Clone)
                .Select(RemoveWorkbookValues)
                .ToList();
        }

        public static IReadOnlyList<MeasurementRule> MergeAcceptedRules(
            IEnumerable<MeasurementRule> acceptedRules,
            IReadOnlyList<MeasurementRule> submittedRules)
        {
            var submitted = submittedRules ?? new List<MeasurementRule>();
            return (acceptedRules ?? Enumerable.Empty<MeasurementRule>())
                .Where(rule => rule != null)
                .Select(rule =>
                {
                    var clone = MeasurementRuleCloner.Clone(rule);
                    var matching = submitted.FirstOrDefault(candidate => SameRule(candidate, clone));
                    if (clone.TemplateDefinition == null && matching?.TemplateDefinition != null)
                    {
                        clone.TemplateDefinition = TemplateDefinitionCloner.Clone(matching.TemplateDefinition);
                    }

                    ApplySubmittedStandardValueMode(clone, matching);

                    return RemoveWorkbookValues(clone);
                })
                .ToList();
        }

        private static MeasurementRule RemoveWorkbookValues(MeasurementRule rule)
        {
            if (HasRange(rule?.StandardValueSource?.Range) ||
                (rule?.RowMappings ?? new List<MeasurementRowMapping>()).Any(mapping => HasRange(mapping?.StandardValueRange)))
            {
                if (HasManualStandardSelection(rule))
                {
                    rule.FixedStandardValue = ResolveFirstManualValue(rule);
                }
                else
                {
                    rule.FixedStandardValue = null;
                    rule.ManualStandardValues = new List<ManualStandardValue>();
                }
            }

            if (HasRange(rule?.RangeSource?.Range) ||
                (rule?.RowMappings ?? new List<MeasurementRowMapping>()).Any(mapping => HasRange(mapping?.RangeValueRange)))
            {
                rule.FixedReferenceRange = null;
            }

            return rule;
        }

        private static void ApplySubmittedStandardValueMode(MeasurementRule target, MeasurementRule submitted)
        {
            if (target == null || submitted == null)
            {
                return;
            }

            target.ManualStandardValues = (submitted.ManualStandardValues ?? new List<ManualStandardValue>())
                .Where(item => item != null)
                .Select(item => new ManualStandardValue { PointIndex = item.PointIndex, Value = item.Value })
                .ToList();
            if (HasManualStandardSelection(submitted))
            {
                target.FixedStandardValue = ResolveFirstManualValue(submitted);
                target.MeasurementLowerBound = submitted.MeasurementLowerBound;
                target.MeasurementUpperBound = submitted.MeasurementUpperBound;
            }
            else
            {
                target.MeasurementLowerBound = null;
                target.MeasurementUpperBound = null;
            }
        }

        private static bool HasManualStandardSelection(MeasurementRule rule)
        {
            return (rule?.ManualStandardValues ?? new List<ManualStandardValue>()).Count > 0;
        }

        private static double? ResolveFirstManualValue(MeasurementRule rule)
        {
            return (rule?.ManualStandardValues ?? new List<ManualStandardValue>())
                .Where(item => item != null && item.Value.HasValue)
                .OrderBy(item => item.PointIndex)
                .Select(item => item.Value)
                .FirstOrDefault();
        }

        private static bool SameRule(MeasurementRule left, MeasurementRule right)
        {
            return string.Equals(
                NormalizeName(left?.FieldAlias ?? left?.FieldName),
                NormalizeName(right?.FieldAlias ?? right?.FieldName),
                StringComparison.Ordinal);
        }

        private static string NormalizeName(string value)
        {
            return new string((value ?? string.Empty)
                .Where(character => char.IsLetterOrDigit(character))
                .ToArray())
                .ToUpperInvariant();
        }

        private static bool HasRange(CellRange range)
        {
            return range != null &&
                range.StartRow > 0 && range.EndRow >= range.StartRow &&
                range.StartColumn > 0 && range.EndColumn >= range.StartColumn;
        }
    }
}
