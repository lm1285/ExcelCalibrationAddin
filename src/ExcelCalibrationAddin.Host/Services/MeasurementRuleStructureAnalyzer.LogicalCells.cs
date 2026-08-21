using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleStructureAnalyzer
    {
        private static List<LogicalCell> GetLogicalCells(SheetSnapshot sheet, CellRange range)
        {
            var logicalCells = new Dictionary<string, LogicalCell>(StringComparer.Ordinal);
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
                var effectiveRange = GetEffectiveRange(cell);
                if (effectiveRange.EndRow < range.StartRow ||
                    effectiveRange.StartRow > range.EndRow ||
                    effectiveRange.EndColumn < range.StartColumn ||
                    effectiveRange.StartColumn > range.EndColumn)
                {
                    continue;
                }

                var anchor = GetEffectiveAnchorCell(sheet, cell);
                if (anchor == null || (string.IsNullOrWhiteSpace(anchor.Text) && string.IsNullOrWhiteSpace(anchor.Formula)))
                {
                    continue;
                }

                var key = $"{effectiveRange.StartRow}:{effectiveRange.StartColumn}:{effectiveRange.EndRow}:{effectiveRange.EndColumn}";
                if (!logicalCells.ContainsKey(key))
                {
                    logicalCells[key] = new LogicalCell
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

        private static CellMeta GetEffectiveAnchorCell(SheetSnapshot sheet, CellMeta cell)
        {
            if (sheet == null || cell == null)
            {
                return null;
            }

            if (cell.MergeRange == null)
            {
                return cell;
            }

            return sheet.Cells.FirstOrDefault(item =>
                       item.Row == cell.MergeRange.StartRow &&
                       item.Column == cell.MergeRange.StartColumn) ??
                   cell;
        }

        private static CellRange GetEffectiveRange(CellMeta cell)
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

    }
}
