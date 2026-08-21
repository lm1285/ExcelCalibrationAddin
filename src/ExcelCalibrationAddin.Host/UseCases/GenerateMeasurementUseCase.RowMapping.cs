using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Generation;
using ExcelCalibrationAddin.Host.Recognition;
using ExcelCalibrationAddin.Host.Services;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public sealed partial class GenerateMeasurementUseCase
    {
        private static DistributionMode ResolveDistributionMode(GenerationConfiguration configuration)
        {
            var value = configuration?.DefaultDistribution;
            if (Enum.TryParse(value, true, out DistributionMode mode))
            {
                return mode;
            }

            return DistributionMode.Normal;
        }

        private static bool HasWritableCells(MeasurementRule rule)
        {
            return rule?.WritableCells != null && rule.WritableCells.Count > 0;
        }

        private static bool HasUsableStandardValue(MeasurementRule rule, WorkbookSnapshot snapshot)
        {
            if (HasManualStandardMode(rule))
            {
                return GetManualStandardValuesByPoint(rule).Count > 0;
            }

            return rule?.FixedStandardValue.HasValue == true ||
                ResolveStandardValuesByRow(snapshot, rule?.StandardValueSource?.Range).Count > 0;
        }

        private static bool HasManualStandardMode(MeasurementRule rule)
        {
            return (rule?.ManualStandardValues ?? Enumerable.Empty<ManualStandardValue>()).Any(item => item != null);
        }

        private static Dictionary<int, double> GetManualStandardValuesByPoint(MeasurementRule rule)
        {
            return (rule?.ManualStandardValues ?? Enumerable.Empty<ManualStandardValue>())
                .Where(item => item != null && item.PointIndex > 0 && item.Value.HasValue)
                .GroupBy(item => item.PointIndex)
                .ToDictionary(group => group.Key, group => group.Last().Value.Value);
        }

        private static bool HasRowMappedStandardValues(
            IReadOnlyDictionary<int, double> standardValuesByRow,
            IReadOnlyList<int> targetRows)
        {
            return standardValuesByRow != null &&
                targetRows != null &&
                standardValuesByRow.Count > 0 &&
                targetRows.Count > 1 &&
                (standardValuesByRow.Count > 1 ||
                 targetRows.Any(row => standardValuesByRow.ContainsKey(row)));
        }

        private static List<CellAddress> CloneCellAddresses(IEnumerable<CellAddress> cells)
        {
            return (cells ?? Enumerable.Empty<CellAddress>())
                .Where(cell => cell != null && cell.Row > 0 && cell.Column > 0)
                .Select(cell => new CellAddress { Row = cell.Row, Column = cell.Column })
                .ToList();
        }

        private static List<CellAddress> BuildContiguousWritableCells(CellRange range, int valueCount)
        {
            var result = new List<CellAddress>();
            if (!GenerationRuleValidator.HasValidRange(range) || valueCount <= 0)
            {
                return result;
            }

            for (var row = range.StartRow; row <= range.EndRow && result.Count < valueCount; row++)
            {
                for (var column = range.StartColumn; column <= range.EndColumn && result.Count < valueCount; column++)
                {
                    result.Add(new CellAddress { Row = row, Column = column });
                }
            }

            return result;
        }

        private static Dictionary<int, double> ResolveStandardValuesByRow(WorkbookSnapshot snapshot, CellRange range)
        {
            return ResolveStandardValuesByRow(snapshot, range, null);
        }

        private static Dictionary<int, double> ResolveStandardValuesByRow(
            WorkbookSnapshot snapshot,
            CellRange range,
            IReadOnlyList<int> targetRows)
        {
            var result = new Dictionary<int, double>();
            if (snapshot == null || range == null)
            {
                return result;
            }

            var sheet = snapshot.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, range.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
            {
                return result;
            }

            var effectiveRange = AlignRangeToTargetRows(range, targetRows);
            foreach (var logicalCell in MergedCellLogicalRangeResolver.GetTextCells(sheet, effectiveRange))
            {
                var value = ResolveNumber(logicalCell.Anchor?.Text);
                if (!value.HasValue || result.ContainsKey(logicalCell.Range.StartRow))
                {
                    continue;
                }

                var startRow = Math.Max(effectiveRange.StartRow, logicalCell.Range.StartRow);
                var endRow = Math.Min(effectiveRange.EndRow, logicalCell.Range.EndRow);
                for (var row = startRow; row <= endRow; row++)
                {
                    if (!result.ContainsKey(row))
                    {
                        result[row] = value.Value;
                    }
                }
            }

            return result;
        }

        private static CellRange AlignRangeToTargetRows(CellRange range, IReadOnlyList<int> targetRows)
        {
            if (range == null || targetRows == null || targetRows.Count <= 1)
            {
                return range;
            }

            var orderedRows = targetRows
                .Where(row => row > 0)
                .Distinct()
                .OrderBy(row => row)
                .ToList();
            if (orderedRows.Count <= 1)
            {
                return range;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = Math.Min(range.StartRow, orderedRows.First()),
                EndRow = Math.Max(range.EndRow, orderedRows.Last()),
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }

        private static bool TryResolveStandardValueForRow(
            MeasurementRule rule,
            IReadOnlyDictionary<int, double> standardValuesByRow,
            int targetRow,
            out double standardValue)
        {
            if (HasManualStandardMode(rule))
            {
                if (standardValuesByRow != null && standardValuesByRow.TryGetValue(targetRow, out standardValue))
                {
                    return true;
                }

                standardValue = 0;
                return false;
            }

            if (standardValuesByRow != null && standardValuesByRow.TryGetValue(targetRow, out standardValue))
            {
                return true;
            }

            var standardRange = rule?.StandardValueSource?.Range;
            if (standardValuesByRow != null &&
                standardValuesByRow.Count > 0 &&
                standardRange != null &&
                targetRow >= standardRange.StartRow &&
                targetRow <= standardRange.EndRow)
            {
                standardValue = 0;
                return false;
            }

            if (rule?.FixedStandardValue.HasValue == true)
            {
                standardValue = rule.FixedStandardValue.Value;
                return true;
            }

            standardValue = 0;
            return false;
        }

        private static double? ResolveNumber(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            foreach (Match match in NumberRegex.Matches(text))
            {
                var contextStart = Math.Max(0, match.Index - 3);
                var contextLength = Math.Min(text.Length - contextStart, match.Length + 6);
                var context = text.Substring(contextStart, contextLength);
                if (context.Contains("%") || context.Contains("％"))
                {
                    continue;
                }

                double value;
                if (double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                {
                    return value;
                }
            }

            return null;
        }

    }
}
