using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropSnapshotProvider
    {
        private void EnsureCalculated()
        {
            try
            {
                const int calculationDone = 0;
                if (SafeToInt(_workbook.Application.CalculationState) != calculationDone)
                {
                    _workbook.Application.Calculate();
                }
            }
            catch
            {
            }
        }

        private static void ApplyMergedRanges(SheetSnapshot sheet, List<CellRange> mergedRanges)
        {
            if (mergedRanges == null || mergedRanges.Count == 0)
            {
                return;
            }

            foreach (var cell in sheet.Cells)
            {
                var mergeRange = FindMergeRange(mergedRanges, cell.Row, cell.Column);
                if (mergeRange == null)
                {
                    continue;
                }

                cell.IsMerged = true;
                cell.MergeRange = mergeRange;
            }
        }

        private static bool ContainsRange(List<CellRange> ranges, CellRange candidate)
        {
            foreach (var existing in ranges)
            {
                if (existing.StartRow == candidate.StartRow &&
                    existing.StartColumn == candidate.StartColumn &&
                    existing.EndRow == candidate.EndRow &&
                    existing.EndColumn == candidate.EndColumn)
                {
                    return true;
                }
            }

            return false;
        }

        private static CellRange FindMergeRange(List<CellRange> mergedRanges, int row, int column)
        {
            foreach (var mergedRange in mergedRanges)
            {
                if (row >= mergedRange.StartRow &&
                    row <= mergedRange.EndRow &&
                    column >= mergedRange.StartColumn &&
                    column <= mergedRange.EndColumn)
                {
                    return mergedRange;
                }
            }

            return null;
        }

        private List<HeaderPath> BuildHeaderPaths(SheetSnapshot sheet)
        {
            var result = new List<HeaderPath>();
            var headerRowCount = Math.Min(HeaderRowsToInspect, InferHeaderRowCount(sheet));
            if (headerRowCount <= 0)
            {
                return result;
            }

            var cellLookup = BuildCellLookup(sheet);
            var maxColumn = 0;
            foreach (var cell in sheet.Cells)
            {
                if (cell.Column > maxColumn)
                {
                    maxColumn = cell.Column;
                }
            }

            for (var column = 1; column <= maxColumn; column++)
            {
                var levels = new List<string>();
                for (var row = 1; row <= headerRowCount; row++)
                {
                    var text = ResolveHeaderText(cellLookup, row, column);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        levels.Add(text.Trim());
                    }
                }

                if (levels.Count > 0)
                {
                    result.Add(new HeaderPath
                    {
                        Column = column,
                        Levels = levels
                    });
                }
            }

            return result;
        }

        private static Dictionary<string, CellMeta> BuildCellLookup(SheetSnapshot sheet)
        {
            var lookup = new Dictionary<string, CellMeta>(StringComparer.Ordinal);
            foreach (var cell in sheet.Cells)
            {
                lookup[BuildCellKey(cell.Row, cell.Column)] = cell;
            }

            return lookup;
        }

        private static string ResolveHeaderText(Dictionary<string, CellMeta> cellLookup, int row, int column)
        {
            if (!cellLookup.TryGetValue(BuildCellKey(row, column), out var cell))
            {
                return string.Empty;
            }

            if (!string.IsNullOrWhiteSpace(cell.Text))
            {
                return cell.Text;
            }

            if (cell.MergeRange != null &&
                cellLookup.TryGetValue(BuildCellKey(cell.MergeRange.StartRow, cell.MergeRange.StartColumn), out var topLeft))
            {
                return topLeft.Text ?? string.Empty;
            }

            return string.Empty;
        }

        private static string BuildCellKey(int row, int column)
        {
            return row + ":" + column;
        }

        private static int InferHeaderRowCount(SheetSnapshot sheet)
        {
            var nonEmptyTopRows = new HashSet<int>();
            foreach (var cell in sheet.Cells)
            {
                if (cell.Row <= 5 && !string.IsNullOrWhiteSpace(cell.Text))
                {
                    nonEmptyTopRows.Add(cell.Row);
                }
            }

            return nonEmptyTopRows.Count == 0 ? 0 : Math.Min(HeaderRowsToInspect, nonEmptyTopRows.Count);
        }

    }
}
