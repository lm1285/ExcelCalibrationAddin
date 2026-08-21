using System;
using System.Globalization;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Recognition
{
    public static class SheetRowContentAnalyzer
    {
        public static bool IsBlankRow(SheetSnapshot sheet, int row)
        {
            return !sheet.Cells.Any(cell =>
                cell.Row == row &&
                CellHasEffectiveContent(sheet, cell));
        }

        public static bool HasDataInRangeRow(SheetSnapshot sheet, int row, int startColumn, int endColumn)
        {
            return sheet.Cells.Any(cell =>
            {
                var effectiveRange = MergedCellLogicalRangeResolver.ResolveEffectiveRange(cell);
                return cell.Row == row &&
                    effectiveRange.StartRow == row &&
                    effectiveRange.StartColumn == cell.Column &&
                    RangesOverlap(effectiveRange.StartColumn, effectiveRange.EndColumn, startColumn, endColumn) &&
                    CellHasEffectiveContent(sheet, cell) &&
                    !IsTrailingNoteRow(sheet, row);
            });
        }

        public static bool HasNumericDataInRangeRow(SheetSnapshot sheet, int row, int startColumn, int endColumn)
        {
            return sheet.Cells.Any(cell =>
            {
                var effectiveRange = MergedCellLogicalRangeResolver.ResolveEffectiveRange(cell);
                return cell.Row == row &&
                    effectiveRange.StartRow == row &&
                    effectiveRange.StartColumn == cell.Column &&
                    RangesOverlap(effectiveRange.StartColumn, effectiveRange.EndColumn, startColumn, endColumn) &&
                    LooksNumeric(MergedCellLogicalRangeResolver.ResolveAnchorCell(sheet, cell)?.Text);
            });
        }

        public static int CountWritableTemplateCellsInRow(SheetSnapshot sheet, int row, int startColumn, int endColumn)
        {
            return sheet.Cells
                .Where(cell =>
                    cell.Row == row &&
                    cell.Column >= startColumn &&
                    cell.Column <= endColumn)
                .Select(cell => new
                {
                    Cell = cell,
                    EffectiveRange = MergedCellLogicalRangeResolver.ResolveEffectiveRange(cell),
                    Anchor = MergedCellLogicalRangeResolver.ResolveAnchorCell(sheet, cell)
                })
                .Where(item =>
                    item.EffectiveRange.StartRow == row &&
                    item.EffectiveRange.StartColumn == item.Cell.Column &&
                    string.IsNullOrWhiteSpace(item.Anchor?.Text) &&
                    string.IsNullOrWhiteSpace(item.Anchor?.Formula))
                .GroupBy(item => new
                {
                    item.EffectiveRange.StartRow,
                    item.EffectiveRange.StartColumn,
                    item.EffectiveRange.EndRow,
                    item.EffectiveRange.EndColumn
                })
                .Count();
        }

        public static bool IsTrailingNoteRow(SheetSnapshot sheet, int row)
        {
            var rowTexts = sheet.Cells
                .Where(cell => cell.Row == row && !string.IsNullOrWhiteSpace(cell.Text))
                .Select(cell => new
                {
                    Column = cell.MergeRange?.StartColumn ?? cell.Column,
                    EndColumn = cell.MergeRange?.EndColumn ?? cell.Column,
                    Text = NormalizeNoteText(cell.Text)
                })
                .OrderBy(cell => cell.Column)
                .ToList();

            if (rowTexts.Count == 0)
            {
                return false;
            }

            var firstTextCell = rowTexts.FirstOrDefault(item => item.Column <= 3 || item.EndColumn >= 1);
            if (firstTextCell == null || firstTextCell.Column > 3)
            {
                return false;
            }

            var firstText = firstTextCell.Text;
            if (LooksLikeNoteText(firstText))
            {
                return true;
            }

            var combinedText = string.Concat(rowTexts.Where(item => item.Column <= 8).Select(item => item.Text));
            return LooksLikeNoteText(combinedText);
        }

        public static bool CellHasEffectiveContent(SheetSnapshot sheet, CellMeta cell)
        {
            if (cell == null)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(cell.Text) || !string.IsNullOrWhiteSpace(cell.Formula))
            {
                return true;
            }

            var anchorCell = MergedCellLogicalRangeResolver.ResolveAnchorCell(sheet, cell);
            return anchorCell != null &&
                (!string.IsNullOrWhiteSpace(anchorCell.Text) || !string.IsNullOrWhiteSpace(anchorCell.Formula));
        }

        public static bool LooksNumeric(string text)
        {
            var value = (text ?? string.Empty).Trim().Replace(",", string.Empty).Replace("%", string.Empty);
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _) ||
                   double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out _);
        }

        public static bool RangesOverlap(int leftStart, int leftEnd, int rightStart, int rightEnd)
        {
            return leftStart <= rightEnd && rightStart <= leftEnd;
        }

        private static string NormalizeNoteText(string text)
        {
            var value = (text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var chars = value
                .Where(ch => !char.IsWhiteSpace(ch) && ch != '\'' && ch != '"' && ch != '\u201C' && ch != '\u201D')
                .ToArray();
            return new string(chars);
        }

        private static bool LooksLikeNoteText(string text)
        {
            var value = NormalizeNoteText(text);
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return value.StartsWith("\u5907\u6CE8", StringComparison.OrdinalIgnoreCase) ||
                value.StartsWith("\u6CE8\u91CA", StringComparison.OrdinalIgnoreCase) ||
                value.StartsWith("\u8BF4\u660E", StringComparison.OrdinalIgnoreCase) ||
                value.StartsWith("\u6CE8\uFF1A", StringComparison.OrdinalIgnoreCase) ||
                value.StartsWith("\u6CE8:", StringComparison.OrdinalIgnoreCase) ||
                value.IndexOf("\u5907\u6CE8", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("\u8868\u793A\u65E0", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("\u4E0D\u6D89\u53CA", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
