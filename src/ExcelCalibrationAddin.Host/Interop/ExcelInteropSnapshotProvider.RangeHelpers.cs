using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropSnapshotProvider
    {
        private static string BuildRangeShape(List<ScanArea> scanAreas, int rowCount, int columnCount)
        {
            if (scanAreas != null && scanAreas.Count > 0)
            {
                return string.Join(",", scanAreas.Select(area => area.Address));
            }

            if (rowCount <= 0 || columnCount <= 0)
            {
                return "A1";
            }

            return $"A1:{ExcelAddressHelper.ToColumnName(columnCount)}{rowCount}";
        }

        private static bool IsValidRange(CellRange range)
        {
            return range != null &&
                !string.IsNullOrWhiteSpace(range.SheetName) &&
                range.StartRow > 0 &&
                range.StartColumn > 0 &&
                range.EndRow >= range.StartRow &&
                range.EndColumn >= range.StartColumn;
        }

        private static List<CellRange> MergeRanges(IReadOnlyList<CellRange> ranges)
        {
            var pending = (ranges ?? Array.Empty<CellRange>())
                .Where(IsValidRange)
                .GroupBy(range => $"{range.StartRow}:{range.StartColumn}:{range.EndRow}:{range.EndColumn}", StringComparer.Ordinal)
                .Select(group => CloneCellRange(group.First()))
                .OrderBy(range => range.StartRow)
                .ThenBy(range => range.StartColumn)
                .ToList();
            var merged = new List<CellRange>();
            foreach (var range in pending)
            {
                var current = range;
                var mergedExisting = true;
                while (mergedExisting)
                {
                    mergedExisting = false;
                    for (var index = 0; index < merged.Count; index++)
                    {
                        if (!CanMergeRanges(merged[index], current))
                        {
                            continue;
                        }

                        current = UnionRanges(merged[index], current);
                        merged.RemoveAt(index);
                        mergedExisting = true;
                        break;
                    }
                }
                merged.Add(current);
            }

            return merged
                .OrderBy(range => range.StartRow)
                .ThenBy(range => range.StartColumn)
                .ToList();
        }

        private static bool CanMergeRanges(CellRange left, CellRange right)
        {
            return left != null && right != null &&
                string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase) &&
                left.StartRow <= right.EndRow + 1 && right.StartRow <= left.EndRow + 1 &&
                left.StartColumn <= right.EndColumn + 1 && right.StartColumn <= left.EndColumn + 1;
        }

        private static CellRange UnionRanges(CellRange left, CellRange right)
        {
            return new CellRange
            {
                SheetName = left.SheetName,
                StartRow = Math.Min(left.StartRow, right.StartRow),
                StartColumn = Math.Min(left.StartColumn, right.StartColumn),
                EndRow = Math.Max(left.EndRow, right.EndRow),
                EndColumn = Math.Max(left.EndColumn, right.EndColumn)
            };
        }

        private static CellRange CloneCellRange(CellRange range)
        {
            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                StartColumn = range.StartColumn,
                EndRow = range.EndRow,
                EndColumn = range.EndColumn
            };
        }

        private static string ToAddress(CellRange range)
        {
            return $"{ExcelAddressHelper.ToColumnName(range.StartColumn)}{range.StartRow}:{ExcelAddressHelper.ToColumnName(range.EndColumn)}{range.EndRow}";
        }

        private static List<ScanArea> ResolveScanAreas(
            dynamic worksheet,
            int usedRangeStartRow,
            int usedRangeStartColumn,
            int rowCount,
            int columnCount)
        {
            var areas = new List<ScanArea>();
            var printAreaText = SafeToString(worksheet.PageSetup?.PrintArea);
            if (!string.IsNullOrWhiteSpace(printAreaText))
            {
                try
                {
                    dynamic printRange = worksheet.Range[printAreaText];
                    foreach (var area in printRange.Areas)
                    {
                        AddScanArea(areas, area);
                    }
                }
                catch
                {
                    areas.Clear();
                }
            }

            if (areas.Count == 0 && rowCount > 0 && columnCount > 0)
            {
                var fallbackRowCount = Math.Min(rowCount, MaxRowsToScan);
                var fallbackColumnCount = Math.Min(columnCount, MaxColumnsToScan);
                FitWithinCellLimit(ref fallbackRowCount, ref fallbackColumnCount);
                areas.Add(new ScanArea
                {
                    StartRow = usedRangeStartRow,
                    StartColumn = usedRangeStartColumn,
                    EndRow = usedRangeStartRow + fallbackRowCount - 1,
                    EndColumn = usedRangeStartColumn + fallbackColumnCount - 1
                });
            }

            return areas
                .Select(TrimArea)
                .Where(area => area.RowCount > 0 && area.ColumnCount > 0)
                .ToList();
        }

        private static void AddScanArea(List<ScanArea> areas, dynamic area)
        {
            var startRow = SafeToInt(area.Row);
            var startColumn = SafeToInt(area.Column);
            var rowCount = SafeToInt(area.Rows.Count);
            var columnCount = SafeToInt(area.Columns.Count);
            if (startRow <= 0 || startColumn <= 0 || rowCount <= 0 || columnCount <= 0)
            {
                return;
            }

            areas.Add(new ScanArea
            {
                StartRow = startRow,
                StartColumn = startColumn,
                EndRow = startRow + rowCount - 1,
                EndColumn = startColumn + columnCount - 1
            });
        }

        private static ScanArea TrimArea(ScanArea area)
        {
            var rowCount = Math.Min(area.RowCount, MaxRowsToScan);
            var columnCount = Math.Min(area.ColumnCount, MaxColumnsToScan);
            FitWithinCellLimit(ref rowCount, ref columnCount);
            return new ScanArea
            {
                StartRow = area.StartRow,
                StartColumn = area.StartColumn,
                EndRow = area.StartRow + rowCount - 1,
                EndColumn = area.StartColumn + columnCount - 1
            };
        }

        private static void FitWithinCellLimit(ref int rowCount, ref int columnCount)
        {
            while (rowCount > 0 && columnCount > 0 && (long)rowCount * columnCount > MaxCellsToScan)
            {
                if (rowCount >= columnCount && rowCount > HeaderRowsToInspect)
                {
                    rowCount--;
                }
                else if (columnCount > 1)
                {
                    columnCount--;
                }
                else
                {
                    rowCount--;
                }
            }
        }

        private static object[,] ReadMatrix(dynamic rawValue, int rowCount, int columnCount)
        {
            var matrix = new object[rowCount + 1, columnCount + 1];
            if (rawValue == null)
            {
                return matrix;
            }

            if (rowCount == 1 && columnCount == 1)
            {
                matrix[1, 1] = rawValue;
                return matrix;
            }

            if (rawValue is object[,])
            {
                var values = (object[,])rawValue;
                for (var row = 1; row <= rowCount; row++)
                {
                    for (var column = 1; column <= columnCount; column++)
                    {
                        matrix[row, column] = values[row, column];
                    }
                }

                return matrix;
            }

            for (var row = 1; row <= rowCount; row++)
            {
                for (var column = 1; column <= columnCount; column++)
                {
                    matrix[row, column] = rawValue;
                }
            }

            return matrix;
        }

        private static List<CellRange> CaptureMergedRangesForCandidates(
            dynamic worksheet,
            IEnumerable<ScanArea> areas,
            IEnumerable<CellAddress> candidates)
        {
            var result = new List<CellRange>();
            foreach (var area in areas ?? Enumerable.Empty<ScanArea>())
            {
                CaptureMergedRangesInArea(worksheet, area, candidates, result);
            }
            return result;
        }

        private static void CaptureMergedRangesInArea(
            dynamic worksheet,
            ScanArea area,
            IEnumerable<CellAddress> candidates,
            List<CellRange> result)
        {
            if (area == null || area.RowCount <= 0 || area.ColumnCount <= 0)
            {
                return;
            }

            dynamic range = worksheet.Range[
                worksheet.Cells[area.StartRow, area.StartColumn],
                worksheet.Cells[area.EndRow, area.EndColumn]];
            var mergeState = ReadMergeState(range);
            if (mergeState == false)
            {
                return;
            }

            foreach (var candidate in (candidates ?? Enumerable.Empty<CellAddress>()).Where(item =>
                item.Row >= area.StartRow &&
                item.Row <= area.EndRow &&
                item.Column >= area.StartColumn &&
                item.Column <= area.EndColumn))
            {
                if (IsCoveredByMergedRange(result, candidate.Row, candidate.Column))
                {
                    continue;
                }

                var mergeRange = ReadMergeRange(worksheet, candidate.Row, candidate.Column);
                if (mergeRange != null && !ContainsRange(result, mergeRange))
                {
                    result.Add(mergeRange);
                }
            }
        }

        private static bool? ReadMergeState(dynamic range)
        {
            try
            {
                var value = range.MergeCells;
                return value == null || value is DBNull ? (bool?)null : Convert.ToBoolean(value);
            }
            catch
            {
                return null;
            }
        }

        private static CellRange ReadMergeRange(dynamic worksheet, int row, int column)
        {
            try
            {
                dynamic cell = worksheet.Cells[row, column];
                if (!SafeToBool(cell.MergeCells))
                {
                    return null;
                }

                dynamic mergeArea = cell.MergeArea;
                var startRow = SafeToInt(mergeArea.Row);
                var startColumn = SafeToInt(mergeArea.Column);
                var rowCount = SafeToInt(mergeArea.Rows.Count);
                var columnCount = SafeToInt(mergeArea.Columns.Count);
                if (startRow <= 0 || startColumn <= 0 || rowCount <= 0 || columnCount <= 0)
                {
                    return null;
                }

                return new CellRange
                {
                    SheetName = SafeToString(worksheet.Name),
                    StartRow = startRow,
                    StartColumn = startColumn,
                    EndRow = startRow + rowCount - 1,
                    EndColumn = startColumn + columnCount - 1
                };
            }
            catch
            {
                return null;
            }
        }

        private static bool IsCoveredByMergedRange(
            IEnumerable<CellRange> mergedRanges,
            int row,
            int column)
        {
            return (mergedRanges ?? Enumerable.Empty<CellRange>()).Any(range =>
                range != null &&
                row >= range.StartRow &&
                row <= range.EndRow &&
                column >= range.StartColumn &&
                column <= range.EndColumn);
        }

    }
}
