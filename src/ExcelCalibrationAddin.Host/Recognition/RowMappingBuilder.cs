using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Recognition
{
    public sealed class RowMappingBuilder
    {
        public void Apply(WorkbookSnapshot snapshot, IReadOnlyList<MeasurementRule> rules)
        {
            foreach (var rule in rules ?? Array.Empty<MeasurementRule>())
            {
                rule.RowMappings = Build(snapshot, rule);
            }
        }

        public List<MeasurementRowMapping> Build(WorkbookSnapshot snapshot, MeasurementRule rule)
        {
            if (snapshot == null || rule?.TargetRange == null)
            {
                return new List<MeasurementRowMapping>();
            }

            var sheet = snapshot.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, rule.TargetRange.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
            {
                return new List<MeasurementRowMapping>();
            }

            var measurementCells = (rule.WritableCells ?? new List<CellAddress>())
                .Where(cell => cell != null)
                .GroupBy(cell => cell.Row)
                .ToDictionary(group => group.Key, group => group.OrderBy(cell => cell.Column).ToList());
            var usesMaximumError = ErrorFormulaClassifier.IsMaximumError(rule);
            var result = new List<MeasurementRowMapping>();
            foreach (var row in measurementCells.Keys.OrderBy(value => value))
            {
                var mapping = new MeasurementRowMapping
                {
                    Row = row,
                    SetpointValueRange = ResolveRangeForRow(sheet, rule.SetpointSource?.Range, row),
                    StandardValueRange = ResolveRangeForRow(sheet, rule.StandardValueSource?.Range, row),
                    MeasurementCells = measurementCells[row],
                    AverageRange = ResolveRangeForRow(sheet, rule.AverageSource?.Range, row),
                    ErrorRange = ResolveRangeForRow(sheet, rule.ErrorSource?.Range, row),
                    TechnicalRequirementRange = ResolveRangeForRow(sheet, rule.MpeSource?.Range, row),
                    RangeValueRange = ResolveRangeForRow(sheet, rule.RangeSource?.Range, row),
                    UncertaintyRange = ResolveRangeForRow(sheet, rule.UncertaintySource?.Range, row),
                    ResultRange = ResolveRangeForRow(sheet, rule.ResultSource?.Range, row)
                };
                var missing = BuildMissingFields(mapping, usesMaximumError);
                mapping.IsComplete = missing.Count == 0;
                mapping.StatusMessage = mapping.IsComplete
                    ? "行级映射完整"
                    : "缺少" + string.Join("、", missing);
                result.Add(mapping);
            }

            return result;
        }

        private static CellRange ResolveRangeForRow(SheetSnapshot sheet, CellRange source, int row)
        {
            if (source == null || row < source.StartRow || row > source.EndRow)
            {
                return null;
            }

            var logical = MergedCellLogicalRangeResolver.GetContentCells(sheet, source)
                .FirstOrDefault(item => row >= item.Range.StartRow && row <= item.Range.EndRow);
            if (logical != null)
            {
                return CloneRange(logical.Range, source.SheetName);
            }

            return new CellRange
            {
                SheetName = source.SheetName,
                StartRow = row,
                EndRow = row,
                StartColumn = source.StartColumn,
                EndColumn = source.EndColumn
            };
        }

        private static List<string> BuildMissingFields(MeasurementRowMapping mapping, bool usesMaximumError)
        {
            var missing = new List<string>();
            if (mapping.StandardValueRange == null) missing.Add("标准值");
            if (mapping.MeasurementCells.Count == 0) missing.Add("测量值");
            if (!usesMaximumError && mapping.ErrorRange == null) missing.Add("误差");
            if (mapping.TechnicalRequirementRange == null) missing.Add("技术要求");
            return missing;
        }

        private static CellRange CloneRange(CellRange range, string fallbackSheetName)
        {
            return new CellRange
            {
                SheetName = string.IsNullOrWhiteSpace(range.SheetName) ? fallbackSheetName : range.SheetName,
                StartRow = range.StartRow,
                EndRow = range.EndRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }
    }
}
