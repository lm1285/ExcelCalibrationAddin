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
        private static int FindBestRepeatabilityDataRow(SheetSnapshot sheet, CellRange range, int endRow)
        {
            var bestRow = 0;
            var bestCount = 0;
            for (var row = range.StartRow; row <= Math.Min(endRow, range.EndRow); row++)
            {
                var count = CountNumericCellsInRow(sheet, row, range.StartColumn, range.EndColumn);
                if (count > bestCount)
                {
                    bestCount = count;
                    bestRow = row;
                }
            }

            return bestRow > 0 ? bestRow : range.StartRow;
        }

        private static int FindFirstRepeatabilityDataRow(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            for (var row = startRow; row <= endRow; row++)
            {
                if (IsMeasurementAttemptSequenceRow(sheet, row, startColumn, endColumn))
                {
                    var nextRow = FindFirstWritableTemplateRow(sheet, row + 1, endRow, startColumn, endColumn);
                    if (nextRow > 0)
                    {
                        return nextRow;
                    }

                    nextRow = FindFirstNumericDataRow(sheet, row + 1, endRow, startColumn, endColumn);
                    if (nextRow > 0)
                    {
                        return nextRow;
                    }

                    continue;
                }

                if (CountNumericCellsInRow(sheet, row, startColumn, endColumn) > 0 ||
                SheetRowContentAnalyzer.CountWritableTemplateCellsInRow(sheet, row, startColumn, endColumn) > 0)
                {
                    return row;
                }
            }

            return 0;
        }

        private static int FindFirstNumericDataRow(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            for (var row = startRow; row <= endRow; row++)
            {
                if (CountNumericCellsInRow(sheet, row, startColumn, endColumn) > 0)
                {
                    return row;
                }
            }

            return 0;
        }

        private static bool IsMeasurementAttemptSequenceRow(SheetSnapshot sheet, int row, int startColumn, int endColumn)
        {
            var parsedCells = sheet.Cells
                .Where(cell =>
                    cell.Row == row &&
                    cell.Column >= startColumn &&
                    cell.Column <= endColumn &&
                    !string.IsNullOrWhiteSpace(cell.Text))
                    .OrderBy(cell => cell.Column)
                .Select(cell => new
                {
                    AttemptIndex = TryParseMeasurementAttemptIndex(cell.Text),
                    IsIgnoredHeader = IsAverageHeader(cell.Text)
                })
                .ToList();

            if (parsedCells.Any(cell => !cell.AttemptIndex.HasValue && !cell.IsIgnoredHeader))
            {
                return false;
            }

            var attempts = parsedCells
                .Where(cell => cell.AttemptIndex.HasValue)
                .Select(cell => cell.AttemptIndex.Value)
                .ToList();

            if (attempts.Count < 2)
            {
                return false;
            }

            return attempts.First() == 1 && attempts.SequenceEqual(Enumerable.Range(1, attempts.Count));
        }

        private static int? TryParseMeasurementAttemptIndex(string text)
        {
            var value = NormalizeHeaderText(text);
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var match = Regex.Match(value, @"^(?:\u6D4B\u91CF|\u8BFB\u6570|\u8BFB\u503C|\u793A\u503C)?(?:\u7B2C)?(\d+)(?:\u6B21|\u56DE|\u70B9|\u7EC4|\u53F7|#)?$", RegexOptions.IgnoreCase);
            if (!match.Success || !int.TryParse(match.Groups[1].Value, out var index))
            {
                return null;
            }

            return index > 0 ? index : (int?)null;
        }

        private static ColumnRun FindNumericRunInRow(
            SheetSnapshot sheet,
            int row,
            int startColumn,
            int endColumn,
            CellRange standardRange,
            CellRange averageRange)
        {
            return FindBestNumericRunInRow(sheet, row, startColumn, endColumn, standardRange, averageRange);
        }

        private static ColumnRun FindBestNumericRunInRow(
            SheetSnapshot sheet,
            int row,
            int startColumn,
            int endColumn,
            CellRange standardRange,
            CellRange averageRange)
        {
            var spans = sheet.Cells
                .Where(cell =>
                    cell.Row == row &&
                    cell.Column >= startColumn &&
                    cell.Column <= endColumn &&
                    SheetRowContentAnalyzer.LooksNumeric(cell.Text) &&
                    !IsColumnInsideRange(cell.Column, standardRange) &&
                    !IsColumnInsideRange(cell.Column, averageRange))
                .Select(cell =>
                {
                    var effectiveRange = MergedCellLogicalRangeResolver.ResolveEffectiveRange(cell);
                    var spanStart = Math.Max(startColumn, effectiveRange.StartColumn);
                    var spanEnd = Math.Min(endColumn, effectiveRange.EndColumn);
                    return new ColumnRun
                    {
                        StartColumn = spanStart,
                        EndColumn = spanEnd,
                        Count = Math.Max(1, spanEnd - spanStart + 1)
                    };
                })
                .GroupBy(span => $"{span.StartColumn}:{span.EndColumn}", StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();

            return BuildAdjacentColumnRuns(spans)
                .OrderByDescending(run => run.Count)
                .ThenBy(run => run.StartColumn)
                .FirstOrDefault();
        }

        private static int CountNumericCellsInRow(SheetSnapshot sheet, int row, int startColumn, int endColumn)
        {
            return sheet.Cells.Count(cell =>
                cell.Row == row &&
                cell.Column >= startColumn &&
                cell.Column <= endColumn &&
                SheetRowContentAnalyzer.LooksNumeric(cell.Text));
        }

        private static int FindRightBoundaryBeforeKnownFields(
            SheetSnapshot sheet,
            int headerRow,
            int startColumn,
            int endRow,
            params CellRange[] knownRanges)
        {
            var knownStart = knownRanges
                .Where(range => range != null && range.StartColumn > startColumn)
                .Select(range => range.StartColumn)
                .DefaultIfEmpty(InferMaxColumn(sheet) + 1)
                .Min();

            return Math.Max(startColumn, knownStart - 1);
        }

        private static bool HasMeasurementParentHeader(SheetSnapshot sheet, int subHeaderRow, int startColumn, int endColumn, int sectionStartRow)
        {
            var searchStartRow = Math.Max(sectionStartRow, subHeaderRow - 3);
            for (var row = subHeaderRow - 1; row >= searchStartRow; row--)
            {
                var parentHeaders = sheet.Cells
                    .Where(cell =>
                        cell.Row == row &&
                        !string.IsNullOrWhiteSpace(cell.Text) &&
                        IsMeasurementParentHeader(cell.Text) &&
                    SheetRowContentAnalyzer.RangesOverlap(
                            cell.MergeRange?.StartColumn ?? cell.Column,
                            cell.MergeRange?.EndColumn ?? cell.Column,
                            startColumn,
                            endColumn))
                    .ToList();

                if (parentHeaders.Count > 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsMeasurementParentHeader(string text)
        {
            return (!LooksLikeWrongFieldHeader(text, MeasurementKeywords) &&
                    MeasurementKeywords.Any(keyword => MatchesKeyword(text, keyword))) ||
                   RepeatedMeasurementParentKeywords.Any(keyword => MatchesKeyword(text, keyword));
        }

        private static bool IsRepeatabilityMeasurementHeader(string text)
        {
            var value = NormalizeHeaderText(text);
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (IsAverageHeader(value) ||
                value.IndexOf("\u8BEF\u5DEE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("\u7ED3\u679C", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("\u7ED3\u8BBA", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            return value.IndexOf("\u6D4B\u91CF\u503C", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("\u6D4B\u91CF\u6B21\u6570", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("\u6D4B\u91CF/\u6B21\u6570", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("\u6D4B\u91CF\u503C/\u6B21\u6570", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("\u6B21\u6570", StringComparison.OrdinalIgnoreCase) >= 0 ||
                MeasurementKeywords.Any(keyword => MatchesKeyword(value, keyword));
        }

    }
}
