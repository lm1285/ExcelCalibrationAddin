using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleParameterResolver
    {
        private static readonly Regex NumberRegex = new Regex(@"[-+]?\d+(\.\d+)?([eE][-+]?\d+)?", RegexOptions.Compiled);
        private static readonly string[] ReferencedMarkers =
        {
            "%FS",
            "FS%",
            "FULLSCALE",
            "FULL SCALE",
            "SPAN",
            "满量程",
            "满量",
            "量程",
            "全量程"
        };

        public IReadOnlyList<MeasurementRule> Apply(WorkbookSnapshot snapshot, IReadOnlyList<MeasurementRule> rules)
        {
            var resolvedRules = (rules ?? Array.Empty<MeasurementRule>())
                .Where(rule => rule != null)
                .Select(MeasurementRuleCloner.Clone)
                .ToList();

            foreach (var rule in resolvedRules)
            {
                ApplyToRule(snapshot, rule);
            }

            return resolvedRules;
        }

        private void ApplyToRule(WorkbookSnapshot snapshot, MeasurementRule rule)
        {
            if (rule == null)
            {
                return;
            }

            if (snapshot == null)
            {
                return;
            }

            RestoreParameterSourcesFromTemplateDefinition(rule);

            var sheetName = rule.TargetRange?.SheetName ??
                            rule.MpeSource?.Range?.SheetName ??
                            rule.RangeSource?.Range?.SheetName;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return;
            }

            var sheet = snapshot.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
            {
                return;
            }

            var manualStandardValue = (rule.ManualStandardValues ?? new List<ManualStandardValue>())
                .Where(item => item != null && item.PointIndex > 0 && item.Value.HasValue)
                .OrderBy(item => item.PointIndex)
                .Select(item => item.Value)
                .FirstOrDefault();
            var standardValue = ResolveStandardValue(sheet, rule.StandardValueSource?.Range);
            if (manualStandardValue.HasValue)
            {
                rule.FixedStandardValue = manualStandardValue.Value;
            }
            else if (standardValue.HasValue)
            {
                rule.FixedStandardValue = standardValue.Value;
            }

            var mpeResolution = ResolveMpe(sheet, rule);
            if (mpeResolution != null)
            {
                rule.FixedMpe = mpeResolution.Mpe;
                rule.FixedNegativeTolerance = mpeResolution.NegativeTolerance;
                rule.FixedPositiveTolerance = mpeResolution.PositiveTolerance;
                rule.RequirementOperator = mpeResolution.RequirementOperator;
                rule.ErrorType = mpeResolution.ErrorType;
                if (rule.MpeSource != null && !string.IsNullOrWhiteSpace(mpeResolution.ValuePattern))
                {
                    rule.MpeSource.ValuePattern = mpeResolution.ValuePattern;
                }

                if (RequiresReferenceRange(rule, mpeResolution.ErrorType))
                {
                    ApplyReferenceRange(sheet, rule, mpeResolution.ReferenceRange);
                }
            }
            else if (RequiresReferenceRange(rule, rule.ErrorType))
            {
                ApplyReferenceRange(sheet, rule, null);
            }
        }

        private static bool RequiresReferenceRange(MeasurementRule rule, ErrorType resolvedErrorType)
        {
            return resolvedErrorType == ErrorType.Referenced ||
                rule?.ErrorFormula?.Scale == ErrorFormulaScale.RelativeToReferenceRange;
        }

        private static void ApplyReferenceRange(
            SheetSnapshot sheet,
            MeasurementRule rule,
            double? resolvedReferenceRange)
        {
            var referenceRange = resolvedReferenceRange;
            if (!referenceRange.HasValue || referenceRange.Value <= 0)
            {
                referenceRange = ResolveReferenceRange(sheet, rule?.RangeSource?.Range);
            }
            if (referenceRange.HasValue && referenceRange.Value > 0)
            {
                rule.FixedReferenceRange = referenceRange.Value;
            }
            else if (rule?.RangeSource?.Range != null)
            {
                rule.FixedReferenceRange = null;
            }
        }

        private static double? ResolveStandardValue(SheetSnapshot sheet, CellRange range)
        {
            if (sheet == null || range == null)
            {
                return null;
            }

            var logicalCells = MergedCellLogicalRangeResolver.GetTextCells(sheet, range);
            foreach (var logicalCell in logicalCells.OrderBy(item => item.Range.StartRow).ThenBy(item => item.Range.StartColumn))
            {
                var candidates = ExtractNumberCandidates(logicalCell.Anchor?.Text);
                var selected = candidates
                    .Where(item => !ContainsPercentToken(item.Context))
                    .DefaultIfEmpty(candidates.FirstOrDefault())
                    .Select(item => (double?)item?.Value)
                    .FirstOrDefault(value => value.HasValue);
                if (selected.HasValue && !double.IsNaN(selected.Value) && !double.IsInfinity(selected.Value))
                {
                    return selected.Value;
                }
            }

            return null;
        }

    }
}
