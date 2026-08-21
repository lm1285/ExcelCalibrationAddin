using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleDraftBuilder
    {
        private static int FindFirstDataRow(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            for (var row = startRow; row <= endRow; row++)
            {
                if (SheetRowContentAnalyzer.HasDataInRangeRow(sheet, row, startColumn, endColumn))
                {
                    return row;
                }
            }

            return 0;
        }

        private static int FindFirstWritableTemplateRow(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            for (var row = startRow; row <= endRow; row++)
            {
                if (IsMeasurementAttemptSequenceRow(sheet, row, startColumn, endColumn))
                {
                    continue;
                }

                var writableCount = SheetRowContentAnalyzer.CountWritableTemplateCellsInRow(sheet, row, startColumn, endColumn);
                if (writableCount > 0)
                {
                    return row;
                }
            }

            return 0;
        }

        private static int FindLastWritableTemplateRow(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            if (startRow <= 0)
            {
                return endRow;
            }

            var lastRow = startRow;
            for (var row = startRow + 1; row <= endRow; row++)
            {
                if (SheetRowContentAnalyzer.CountWritableTemplateCellsInRow(sheet, row, startColumn, endColumn) <= 0)
                {
                    break;
                }

                lastRow = row;
            }

            return lastRow;
        }

        private static int TrimMeasurementEndColumnBySiblingHeaders(SheetSnapshot sheet, int headerBottomRow, int startColumn, int endColumn)
        {
            var stopColumn = sheet.Cells
                .Where(cell =>
                    cell.Row >= headerBottomRow &&
                    cell.Row <= headerBottomRow + 2 &&
                    cell.Column > startColumn &&
                    cell.Column <= endColumn &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    (IsAverageHeader(cell.Text) ||
                     ErrorKeywords.Any(keyword => MatchesKeyword(cell.Text, keyword)) ||
                     TechnicalKeywords.Any(keyword => MatchesKeyword(cell.Text, keyword)) ||
                     ResultKeywords.Any(keyword => MatchesKeyword(cell.Text, keyword))))
                .Select(cell => cell.MergeRange?.StartColumn ?? cell.Column)
                .DefaultIfEmpty(0)
                .Min();

            return stopColumn > startColumn ? Math.Min(endColumn, stopColumn - 1) : endColumn;
        }

        private static int? FindFirstFormulaRow(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            for (var row = startRow; row <= endRow; row++)
            {
                var hasFormula = sheet.Cells.Any(cell =>
                    cell.Row == row &&
                    cell.Column >= startColumn &&
                    cell.Column <= endColumn &&
                    !string.IsNullOrWhiteSpace(cell.Formula));

                if (hasFormula)
                {
                    return row;
                }
            }

            return null;
        }

        private static bool HasSufficientDataBelow(SheetSnapshot sheet, int headerBottomRow, int endRow, int startColumn, int endColumn)
        {
            return CountDataCells(sheet, headerBottomRow + 1, endRow, startColumn, endColumn) >= 2;
        }

        private static int CountDataCells(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            return sheet.Cells.Count(cell =>
                cell.Row >= startRow &&
                cell.Row <= endRow &&
                cell.Column >= startColumn &&
                cell.Column <= endColumn &&
                SheetRowContentAnalyzer.CellHasEffectiveContent(sheet, cell));
        }

        private static int CountFormulaCells(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            return sheet.Cells.Count(cell =>
                cell.Row >= startRow &&
                cell.Row <= endRow &&
                cell.Column >= startColumn &&
                cell.Column <= endColumn &&
                !string.IsNullOrWhiteSpace(cell.Formula));
        }

        private static int CountPeerHeaders(SheetSnapshot sheet, int row, int sectionStartRow, int sectionEndRow)
        {
            return sheet.Cells.Count(cell =>
                cell.Row == row &&
                cell.Row >= sectionStartRow &&
                cell.Row <= Math.Min(sectionEndRow, sectionStartRow + 6) &&
                !string.IsNullOrWhiteSpace(cell.Text) &&
                !LooksLikeSectionTitle(cell.Text));
        }

    }
}
