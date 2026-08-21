using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Recognition
{
    public static class MergedCellLogicalRangeResolver
    {
        private static readonly ConditionalWeakTable<SheetSnapshot, CellIndex> CellIndexes =
            new ConditionalWeakTable<SheetSnapshot, CellIndex>();

        public static List<LogicalCellRange> GetTextCells(SheetSnapshot sheet, CellRange range)
        {
            return GetCells(sheet, range, anchor => !string.IsNullOrWhiteSpace(anchor?.Text));
        }

        public static List<LogicalCellRange> GetContentCells(SheetSnapshot sheet, CellRange range)
        {
            return GetCells(sheet, range, anchor =>
                !string.IsNullOrWhiteSpace(anchor?.Text) ||
                !string.IsNullOrWhiteSpace(anchor?.Formula));
        }

        public static CellMeta ResolveAnchorCell(SheetSnapshot sheet, CellMeta cell)
        {
            if (sheet == null || cell == null)
            {
                return null;
            }

            if (cell.MergeRange == null)
            {
                return cell;
            }

            var index = CellIndexes.GetValue(sheet, BuildCellIndex);
            return index.Find(cell.MergeRange.StartRow, cell.MergeRange.StartColumn) ?? cell;
        }

        public static CellRange ResolveEffectiveRange(CellMeta cell)
        {
            if (cell?.MergeRange != null)
            {
                return CloneRange(cell.MergeRange);
            }

            return new CellRange
            {
                SheetName = string.Empty,
                StartRow = cell?.Row ?? 0,
                EndRow = cell?.Row ?? 0,
                StartColumn = cell?.Column ?? 0,
                EndColumn = cell?.Column ?? 0
            };
        }

        private static List<LogicalCellRange> GetCells(
            SheetSnapshot sheet,
            CellRange range,
            Func<CellMeta, bool> includeAnchor)
        {
            var logicalCells = new Dictionary<string, LogicalCellRange>(StringComparer.Ordinal);
            if (sheet == null || range == null)
            {
                return logicalCells.Values.ToList();
            }

            foreach (var cell in sheet.Cells.Where(cell =>
                         cell.Row >= range.StartRow &&
                         cell.Row <= range.EndRow &&
                         cell.Column >= range.StartColumn &&
                         cell.Column <= range.EndColumn))
            {
                var effectiveRange = ResolveEffectiveRange(cell);
                if (effectiveRange.EndRow < range.StartRow ||
                    effectiveRange.StartRow > range.EndRow ||
                    effectiveRange.EndColumn < range.StartColumn ||
                    effectiveRange.StartColumn > range.EndColumn)
                {
                    continue;
                }

                var anchor = ResolveAnchorCell(sheet, cell);
                if (!includeAnchor(anchor))
                {
                    continue;
                }

                var key = BuildRangeKey(effectiveRange);
                if (!logicalCells.ContainsKey(key))
                {
                    logicalCells[key] = new LogicalCellRange
                    {
                        Anchor = anchor,
                        Range = effectiveRange
                    };
                }
            }

            return logicalCells.Values
                .OrderBy(item => item.Range.StartRow)
                .ThenBy(item => item.Range.StartColumn)
                .ToList();
        }

        private static CellRange CloneRange(CellRange range)
        {
            if (range == null)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                EndRow = range.EndRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }

        private static string BuildRangeKey(CellRange range)
        {
            return $"{range.StartRow}:{range.StartColumn}:{range.EndRow}:{range.EndColumn}";
        }

        private static CellIndex BuildCellIndex(SheetSnapshot sheet)
        {
            return new CellIndex(sheet?.Cells ?? new List<CellMeta>());
        }

        private sealed class CellIndex
        {
            private readonly Dictionary<long, CellMeta> _cells;

            public CellIndex(IEnumerable<CellMeta> cells)
            {
                _cells = (cells ?? Enumerable.Empty<CellMeta>())
                    .Where(cell => cell != null)
                    .GroupBy(cell => BuildCellKey(cell.Row, cell.Column))
                    .ToDictionary(group => group.Key, group => group.Last());
            }

            public CellMeta Find(int row, int column)
            {
                return _cells.TryGetValue(BuildCellKey(row, column), out var cell)
                    ? cell
                    : null;
            }

            private static long BuildCellKey(int row, int column)
            {
                return ((long)row << 32) | (uint)column;
            }
        }
    }
}
