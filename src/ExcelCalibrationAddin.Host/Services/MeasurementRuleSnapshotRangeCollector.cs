using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Generation;

namespace ExcelCalibrationAddin.Host.Services
{
    internal static class MeasurementRuleSnapshotRangeCollector
    {
        public static List<CellRange> Collect(IEnumerable<MeasurementRule> rules)
        {
            var ranges = new List<CellRange>();
            foreach (var rule in (rules ?? Enumerable.Empty<MeasurementRule>()).Where(item => item != null))
            {
                Add(ranges, rule.TargetRange);
                Add(ranges, rule.StandardValueSource?.Range);
                Add(ranges, rule.AverageSource?.Range);
                Add(ranges, rule.ErrorSource?.Range);
                Add(ranges, rule.MpeSource?.Range);
                Add(ranges, rule.RangeSource?.Range);
                Add(ranges, rule.UncertaintySource?.Range);
                Add(ranges, rule.ResultSource?.Range);
                AddHeaderContext(ranges, rule.ErrorSource?.Range);
                AddHeaderContext(ranges, rule.MpeSource?.Range);

                foreach (var dependency in rule.ErrorFormula?.DependencyRanges ?? new List<CellRange>())
                {
                    Add(ranges, dependency);
                }

                var sheetName = rule.TargetRange?.SheetName ??
                                rule.ErrorSource?.Range?.SheetName ??
                                rule.MpeSource?.Range?.SheetName;
                foreach (var constraint in rule.AdditionalJudgementConstraints ?? new List<MeasurementJudgementConstraint>())
                {
                    Add(ranges, constraint?.ErrorSource?.Range);
                    Add(ranges, constraint?.MpeSource?.Range);
                    Add(ranges, constraint?.ResultSource?.Range);
                    AddHeaderContext(ranges, constraint?.ErrorSource?.Range);
                    AddHeaderContext(ranges, constraint?.MpeSource?.Range);
                    foreach (var formula in new[]
                    {
                        constraint?.ErrorFormula?.Formula,
                        constraint?.ErrorFormula?.TechnicalRequirementFormula,
                        constraint?.ErrorFormula?.ResultFormula
                    })
                    {
                        foreach (var referenced in TemplateFormulaParser.ExtractReferencedRanges(formula, sheetName))
                        {
                            Add(ranges, referenced);
                        }
                    }
                }
                foreach (var formula in new[]
                {
                    rule.ErrorFormula?.Formula,
                    rule.ErrorFormula?.TechnicalRequirementFormula,
                    rule.ErrorFormula?.AverageFormula,
                    rule.ErrorFormula?.ResultFormula
                })
                {
                    foreach (var referenced in TemplateFormulaParser.ExtractReferencedRanges(formula, sheetName))
                    {
                        Add(ranges, referenced);
                    }
                }
            }

            return ranges;
        }

        private static void AddHeaderContext(ICollection<CellRange> ranges, CellRange range)
        {
            if (!GenerationRuleValidator.HasValidRange(range))
            {
                return;
            }

            Add(ranges, new CellRange
            {
                SheetName = range.SheetName,
                StartRow = Math.Max(1, range.StartRow - 4),
                EndRow = Math.Max(1, range.StartRow - 1),
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            });
        }

        private static void Add(ICollection<CellRange> ranges, CellRange range)
        {
            if (GenerationRuleValidator.HasValidRange(range))
            {
                ranges.Add(range);
            }
        }
    }
}
