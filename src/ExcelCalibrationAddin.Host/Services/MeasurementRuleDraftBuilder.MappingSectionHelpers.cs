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
        private static int InferMaxColumn(SheetSnapshot sheet)
        {
            return sheet.Cells.Count == 0 ? 1 : sheet.Cells.Max(item => item.Column);
        }

        private static int TrimTrailingNoteRows(SheetSnapshot sheet, int startRow, int endRow)
        {
            var currentEndRow = Math.Max(startRow, endRow);
            while (currentEndRow > startRow && (SheetRowContentAnalyzer.IsTrailingNoteRow(sheet, currentEndRow) || SheetRowContentAnalyzer.IsBlankRow(sheet, currentEndRow)))
            {
                currentEndRow--;
            }

            return currentEndRow;
        }

        private static int FindSectionEndRow(SheetSnapshot sheet, int startRow, int fallbackEndRow)
        {
            if (sheet == null)
            {
                return fallbackEndRow;
            }

            var nextSectionRow = sheet.Cells
                .Where(cell =>
                    cell.Row > startRow &&
                    cell.Row <= fallbackEndRow &&
                    cell.Column <= 2 &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    LooksLikeSectionTitle(cell.Text))
                .Select(cell => cell.Row)
                .DefaultIfEmpty(0)
                .Min();

            return nextSectionRow > 0
                ? Math.Max(startRow, nextSectionRow - 1)
                : fallbackEndRow;
        }

        private static CellRange FindDataRange(SheetSnapshot sheet, int startRow, int endRow, params string[] keywords)
        {
            var headerCandidates = sheet.Cells
                .Where(cell =>
                    cell.Row >= startRow &&
                    cell.Row <= Math.Min(endRow, startRow + 6) &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    !LooksLikeSectionTitle(cell.Text) &&
                    !LooksLikeWrongFieldHeader(cell.Text, keywords) &&
                    keywords.Any(keyword => MatchesKeyword(cell.Text, keyword)))
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();

            if (headerCandidates.Count == 0)
            {
                return null;
            }

            var selectedHeader = SelectBestHeaderCandidate(sheet, startRow, endRow, keywords, headerCandidates);
            if (selectedHeader == null)
            {
                return null;
            }

            var effectiveRange = RefineMeasurementColumns(
                sheet,
                selectedHeader.HeaderBottomRow,
                endRow,
                selectedHeader.StartColumn,
                selectedHeader.EndColumn,
                keywords);
            var columnStart = effectiveRange.StartColumn;
            var columnEnd = effectiveRange.EndColumn;
            var headerBottomRow = effectiveRange.HeaderBottomRow;
            if (IsMeasurementSearch(keywords))
            {
                columnEnd = TrimMeasurementEndColumnBySiblingHeaders(sheet, selectedHeader.HeaderBottomRow, columnStart, columnEnd);
            }

            var dataStartRow = FindFirstDataRow(sheet, headerBottomRow + 1, endRow, columnStart, columnEnd);

            var dataEndRow = endRow;
            if (dataStartRow <= 0 && IsMeasurementSearch(keywords))
            {
                dataStartRow = FindFirstWritableTemplateRow(sheet, headerBottomRow + 1, endRow, columnStart, columnEnd);
                dataEndRow = FindLastWritableTemplateRow(sheet, dataStartRow, endRow, columnStart, columnEnd);
            }

            if (dataStartRow <= 0 && !HasSufficientDataBelow(sheet, headerBottomRow, endRow, columnStart, columnEnd))
            {
                return null;
            }

            if (dataStartRow <= 0)
            {
                dataStartRow = Math.Min(endRow, headerBottomRow + 1);
            }

            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = dataStartRow,
                EndRow = dataEndRow,
                StartColumn = columnStart,
                EndColumn = columnEnd
            };
        }

        private static HeaderBand InferHeaderBand(int sectionStartRow, params CellRange[] ranges)
        {
            var firstDataRow = ranges
                .Where(range => range != null)
                .Select(range => range.StartRow)
                .DefaultIfEmpty(sectionStartRow + 2)
                .Min();
            var bottomRow = Math.Max(sectionStartRow, firstDataRow - 1);
            var topRow = Math.Max(sectionStartRow, bottomRow - 1);
            return new HeaderBand
            {
                TopRow = topRow,
                BottomRow = bottomRow
            };
        }

        private static CellRange InferRangeFromLayout(
            SheetSnapshot sheet,
            HeaderBand headerBand,
            int endRow,
            string[] keywords,
            params CellRange[] anchorRanges)
        {
            if (headerBand == null)
            {
                return null;
            }

            var candidates = sheet.Cells
                .Where(cell =>
                    cell.Row >= headerBand.TopRow &&
                    cell.Row <= headerBand.BottomRow &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    !LooksLikeSectionTitle(cell.Text) &&
                    !LooksLikeWrongFieldHeader(cell.Text, keywords) &&
                    !anchorRanges.Any(range => IsColumnInsideRange(cell.MergeRange?.StartColumn ?? cell.Column, range)) &&
                    keywords.Any(keyword => MatchesKeyword(cell.Text, keyword)))
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => ScoreAnchorDistance(cell, anchorRanges))
                .ThenBy(cell => cell.Column)
                .ToList();

            if (candidates.Count == 0)
            {
                return null;
            }

            var best = candidates.First();
            var columnStart = best.MergeRange?.StartColumn ?? best.Column;
            var columnEnd = best.MergeRange?.EndColumn ?? best.Column;
            var headerBottomRow = Math.Max(headerBand.BottomRow, best.MergeRange?.EndRow ?? best.Row);
            var dataStartRow = FindFirstDataRow(sheet, headerBottomRow + 1, endRow, columnStart, columnEnd);
            if (dataStartRow <= 0)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = dataStartRow,
                EndRow = endRow,
                StartColumn = columnStart,
                EndColumn = columnEnd
            };
        }

        private static CellRange FindInlineParameterValueRange(SheetSnapshot sheet, int startRow, int endRow, params string[] keywords)
        {
            var searchEndRow = Math.Min(endRow, startRow + 12);
            var labelCells = sheet.Cells
                .Where(cell =>
                    cell.Row >= startRow &&
                    cell.Row <= searchEndRow &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    keywords.Any(keyword => MatchesKeyword(cell.Text, keyword)))
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();

            foreach (var labelCell in labelCells)
            {
                var text = (labelCell.Text ?? string.Empty).Trim();
                if (HasInlineParameterValue(text, keywords))
                {
                    return BuildSingleCellRange(sheet.Name, labelCell);
                }

                var valueCell = FindRightSideValueCell(sheet, labelCell, endRow);
                if (valueCell != null)
                {
                    return BuildSingleCellRange(sheet.Name, valueCell);
                }
            }

            return null;
        }

        private static CellRange FindInlineRangeValueRange(SheetSnapshot sheet, int startRow, int endRow)
        {
            if (sheet == null)
            {
                return null;
            }

            var searchStartRow = Math.Max(1, startRow - 18);
            var searchEndRow = Math.Min(endRow, startRow + 12);
            var labels = sheet.Cells
                .Where(cell =>
                    cell.Row >= searchStartRow &&
                    cell.Row <= searchEndRow &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    IsExplicitRangeLabel(cell.Text))
                .OrderByDescending(cell => cell.Row <= startRow)
                .ThenByDescending(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();

            foreach (var label in labels)
            {
                var labelEndColumn = label.MergeRange?.EndColumn ?? label.Column;
                if (HasInlineParameterValue(label.Text, RangeKeywords))
                {
                    return BuildSingleCellRange(sheet.Name, label);
                }

                var numericCells = sheet.Cells
                    .Where(cell =>
                        cell.Row == label.Row &&
                        cell.Column > labelEndColumn &&
                        cell.Column <= labelEndColumn + 12 &&
                        SheetRowContentAnalyzer.LooksNumeric(cell.Text))
                    .Select(cell => MergedCellLogicalRangeResolver.ResolveEffectiveRange(cell))
                    .ToList();
                if (numericCells.Count == 0)
                {
                    continue;
                }

                return new CellRange
                {
                    SheetName = sheet.Name,
                    StartRow = numericCells.Min(range => range.StartRow),
                    EndRow = numericCells.Max(range => range.EndRow),
                    StartColumn = numericCells.Min(range => range.StartColumn),
                    EndColumn = numericCells.Max(range => range.EndColumn)
                };
            }

            return null;
        }

        private static bool IsExplicitRangeLabel(string text)
        {
            var normalized = NormalizeHeaderText(text);
            return RangeKeywords
                .Where(keyword => !string.Equals(keyword, "FS", StringComparison.OrdinalIgnoreCase) &&
                                  !string.Equals(keyword, "Full Scale", StringComparison.OrdinalIgnoreCase) &&
                                  !string.Equals(keyword, "Span", StringComparison.OrdinalIgnoreCase))
                .Any(keyword => string.Equals(normalized, NormalizeHeaderText(keyword), StringComparison.OrdinalIgnoreCase));
        }

        private static CellRange FindInlineParameterRegion(SheetSnapshot sheet, int startRow, int endRow, params string[] keywords)
        {
            var searchEndRow = Math.Min(endRow, startRow + 12);
            var printEndColumn = InferPrintEndColumn(sheet);
            var labelCells = sheet.Cells
                .Where(cell =>
                    cell.Row >= startRow &&
                    cell.Row <= searchEndRow &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    keywords.Any(keyword => MatchesKeyword(cell.Text, keyword)))
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();

            foreach (var labelCell in labelCells)
            {
                var labelEndColumn = labelCell.MergeRange?.EndColumn ?? labelCell.Column;
                if (labelEndColumn >= printEndColumn)
                {
                    continue;
                }

                var regionStartColumn = labelEndColumn + 1;
                var regionEndColumn = printEndColumn;
                var valueColumns = sheet.Cells
                    .Where(cell =>
                        cell.Row == labelCell.Row &&
                        cell.Column >= regionStartColumn &&
                        cell.Column <= regionEndColumn &&
                        !string.IsNullOrWhiteSpace(cell.Text))
                    .Select(cell => cell.Column)
                    .ToList();

                if (valueColumns.Count == 0 && !HasSufficientDataBelow(sheet, labelCell.Row, Math.Min(endRow, labelCell.Row + 2), regionStartColumn, regionEndColumn))
                {
                    continue;
                }

                return new CellRange
                {
                    SheetName = sheet.Name,
                    StartRow = labelCell.Row,
                    EndRow = labelCell.Row,
                    StartColumn = valueColumns.Count == 0 ? regionStartColumn : valueColumns.Min(),
                    EndColumn = regionEndColumn
                };
            }

            return null;
        }

    }
}
