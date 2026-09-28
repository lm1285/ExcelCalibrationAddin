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
            if (rule == null)
            {
                return null;
            }

            // A saved template is a location/relationship description.  These
            // fields are derived from the workbook snapshot and must be
            // rebuilt on every generation, otherwise a changed formula or
            // requirement silently reuses the old MPE/unit interpretation.
            rule.FixedMpe = null;
            rule.FixedNegativeTolerance = null;
            rule.FixedPositiveTolerance = null;
            rule.FixedReferenceRange = null;
            rule.RequirementOperator = TechnicalRequirementOperator.None;
            rule.ErrorFormula = null;
            rule.MpeSource = ClearDynamicPattern(rule.MpeSource);
            rule.RangeSource = ClearDynamicPattern(rule.RangeSource);

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

            foreach (var constraint in rule.AdditionalJudgementConstraints ?? new List<MeasurementJudgementConstraint>())
            {
                if (constraint == null) continue;
                constraint.FixedMpe = null;
                constraint.FixedNegativeTolerance = null;
                constraint.FixedPositiveTolerance = null;
                constraint.RequirementOperator = TechnicalRequirementOperator.None;
                constraint.ErrorFormula = null;
                constraint.MpeSource = ClearDynamicPattern(constraint.MpeSource);
            }

            foreach (var mapping in rule.RowMappings ?? new List<MeasurementRowMapping>())
            {
                foreach (var constraint in mapping?.AdditionalJudgementConstraints ?? new List<MeasurementJudgementConstraint>())
                {
                    if (constraint == null) continue;
                    constraint.FixedMpe = null;
                    constraint.FixedNegativeTolerance = null;
                    constraint.FixedPositiveTolerance = null;
                    constraint.RequirementOperator = TechnicalRequirementOperator.None;
                    constraint.ErrorFormula = null;
                    constraint.MpeSource = ClearDynamicPattern(constraint.MpeSource);
                }
            }

            // Keep only location and role metadata in the template definition.
            // Header text, unit strings, requirement literals, and formula text
            // are all workbook state and must be reread from the live workbook.
            foreach (var region in rule.TemplateDefinition?.Regions ?? new List<TemplateRegionDefinition>())
            {
                if (region == null) continue;
                region.Unit = string.Empty;
                region.Units = new List<string>();
                region.HeaderPath = new List<string>();
                region.NumberFormat = string.Empty;
                region.Formula = null;
                region.FormulaVariants = new List<TemplateFormulaDefinition>();
                region.RequirementValues = new List<TemplateRequirementValue>();
                region.OperatorRange = null;
                region.ValueRange = null;
            }
            foreach (var header in rule.TemplateDefinition?.Headers ?? new List<TemplateHeaderDefinition>())
            {
                if (header == null) continue;
                header.Text = string.Empty;
                header.Unit = string.Empty;
                header.NumberFormat = string.Empty;
            }

            return rule;
        }

        private static ParameterSource ClearDynamicPattern(ParameterSource source)
        {
            if (source != null)
            {
                source.ValuePattern = string.Empty;
            }

            return source;
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
