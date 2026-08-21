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
        private static RefinedRange RefineMeasurementColumns(
            SheetSnapshot sheet,
            int headerBottomRow,
            int endRow,
            int columnStart,
            int columnEnd,
            string[] keywords)
        {
            var isMeasurementRange = IsMeasurementSearch(keywords);

            if (!isMeasurementRange)
            {
                return new RefinedRange
                {
                    StartColumn = columnStart,
                    EndColumn = columnEnd,
                    HeaderBottomRow = headerBottomRow
                };
            }

            var subHeaderSearchEndColumn = IsMeasurementSearch(keywords)
                ? ExpandMeasurementSubHeaderSearchEndColumn(sheet, headerBottomRow, columnStart, columnEnd)
                : columnEnd;
            var subHeaderStartRow = IsMeasurementSearch(keywords)
                ? Math.Max(1, headerBottomRow - 1)
                : headerBottomRow + 1;
            var subHeaderRow = FindSubHeaderRow(sheet, subHeaderStartRow, endRow, columnStart, subHeaderSearchEndColumn);
            if (subHeaderRow <= 0)
            {
                return new RefinedRange
                {
                    StartColumn = columnStart,
                    EndColumn = columnEnd,
                    HeaderBottomRow = headerBottomRow
                };
            }

            var measurementSpans = GetMeasurementSubHeaderSpans(sheet, subHeaderRow, columnStart, subHeaderSearchEndColumn);
            if (measurementSpans.Count == 0)
            {
                return new RefinedRange
                {
                    StartColumn = columnStart,
                    EndColumn = columnEnd,
                    HeaderBottomRow = subHeaderRow
                };
            }

            var measurementRun = MergeAdjacentColumnSpans(measurementSpans);
            return new RefinedRange
            {
                StartColumn = measurementRun.StartColumn,
                EndColumn = measurementRun.EndColumn,
                HeaderBottomRow = subHeaderRow
            };
        }

        private static int ExpandMeasurementSubHeaderSearchEndColumn(SheetSnapshot sheet, int headerBottomRow, int columnStart, int columnEnd)
        {
            var maxColumn = InferMaxColumn(sheet);
            var searchEndColumn = Math.Min(maxColumn, Math.Max(columnEnd, columnStart + 12));
            var boundary = sheet.Cells
                .Where(cell =>
                    cell.Row >= headerBottomRow &&
                    cell.Row <= headerBottomRow + 3 &&
                    cell.Column > columnStart &&
                    cell.Column <= searchEndColumn &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    (IsAverageHeader(cell.Text) ||
                     ErrorKeywords.Any(keyword => MatchesKeyword(cell.Text, keyword)) ||
                     TechnicalKeywords.Any(keyword => MatchesKeyword(cell.Text, keyword)) ||
                     ResultKeywords.Any(keyword => MatchesKeyword(cell.Text, keyword))))
                .Select(cell => cell.MergeRange?.StartColumn ?? cell.Column)
                .DefaultIfEmpty(0)
                .Min();

            return boundary > columnStart ? boundary - 1 : searchEndColumn;
        }

        private static List<ColumnRun> GetMeasurementSubHeaderSpans(SheetSnapshot sheet, int row, int startColumn, int endColumn)
        {
            return sheet.Cells
                .Where(cell =>
                    cell.Row == row &&
                    cell.Column >= startColumn &&
                    cell.Column <= endColumn &&
                    IsMeasurementSubHeader(cell.Text))
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
                .OrderBy(span => span.StartColumn)
                .ThenBy(span => span.EndColumn)
                .ToList();
        }

        private static ColumnRun MergeAdjacentColumnSpans(IReadOnlyList<ColumnRun> spans)
        {
            if (spans == null || spans.Count == 0)
            {
                return null;
            }

            return new ColumnRun
            {
                StartColumn = spans.Min(span => span.StartColumn),
                EndColumn = spans.Max(span => span.EndColumn),
                Count = spans.Sum(span => Math.Max(1, span.Count))
            };
        }

        private static List<ColumnRun> BuildAdjacentColumnRuns(IReadOnlyList<ColumnRun> spans)
        {
            var result = new List<ColumnRun>();
            if (spans == null || spans.Count == 0)
            {
                return result;
            }

            ColumnRun current = null;
            foreach (var span in spans.OrderBy(item => item.StartColumn).ThenBy(item => item.EndColumn))
            {
                if (current == null || span.StartColumn > current.EndColumn + 1)
                {
                    current = new ColumnRun
                    {
                        StartColumn = span.StartColumn,
                        EndColumn = span.EndColumn,
                        Count = Math.Max(1, span.Count)
                    };
                    result.Add(current);
                    continue;
                }

                current.EndColumn = Math.Max(current.EndColumn, span.EndColumn);
                current.Count += Math.Max(1, span.Count);
            }

            return result;
        }

    }
}
