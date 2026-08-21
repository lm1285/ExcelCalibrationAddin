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
        private static CellRange InferMeasurementRangeFromNumericHeaders(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            CellRange standardRange,
            CellRange averageRange)
        {
            var searchEndRow = Math.Min(endRow, startRow + 8);
            foreach (var anchorRange in new[] { standardRange, averageRange }.Where(range => range?.StartRow > startRow))
            {
                searchEndRow = Math.Min(searchEndRow, anchorRange.StartRow - 1);
            }

            NumericHeaderRun best = null;

            for (var row = startRow; row <= searchEndRow; row++)
            {
                var headerSpans = GetMeasurementSubHeaderSpans(sheet, row, 1, InferMaxColumn(sheet))
                .Where(span => !SheetRowContentAnalyzer.RangesOverlap(span.StartColumn, span.EndColumn, standardRange?.StartColumn ?? 0, standardRange?.EndColumn ?? 0))
                    .ToList();

                foreach (var run in BuildAdjacentColumnRuns(headerSpans))
                {
                    if (run.Count == 0)
                    {
                        continue;
                    }

                    if (!HasMeasurementParentHeader(sheet, row, run.StartColumn, run.EndColumn, startRow))
                    {
                        continue;
                    }

                    var dataStartRow = FindFirstDataRow(sheet, row + 1, endRow, run.StartColumn, run.EndColumn);
                    if (dataStartRow <= 0)
                    {
                        continue;
                    }

                    var dataCells = CountDataCells(sheet, dataStartRow, endRow, run.StartColumn, run.EndColumn);
                    var score = dataCells + run.Count * 3 - Math.Max(0, row - startRow);
                    if (best == null || score > best.Score)
                    {
                        best = new NumericHeaderRun
                        {
                            HeaderRow = row,
                            StartColumn = run.StartColumn,
                            EndColumn = run.EndColumn,
                            DataStartRow = dataStartRow,
                            Score = score
                        };
                    }
                }
            }

            if (best == null)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = best.DataStartRow,
                EndRow = endRow,
                StartColumn = best.StartColumn,
                EndColumn = best.EndColumn
            };
        }

        private static CellRange RefineRangeFromSelectedMeasurementSubHeaders(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            CellRange range)
        {
            if (range == null)
            {
                return null;
            }

            var headerRow = FindSubHeaderRow(sheet, startRow, Math.Max(startRow, range.StartRow - 1), range.StartColumn, range.EndColumn);
            if (headerRow <= 0)
            {
                return range;
            }

            var selectedSubHeaderSpans = GetMeasurementSubHeaderSpans(sheet, headerRow, range.StartColumn, range.EndColumn);
            if (selectedSubHeaderSpans.Count == 0)
            {
                return range;
            }

            var runs = BuildAdjacentColumnRuns(selectedSubHeaderSpans);
            var bestRun = runs
                .Where(run => run.Count >= 1 && HasMeasurementParentHeader(sheet, headerRow, run.StartColumn, run.EndColumn, startRow))
                .OrderByDescending(run => CountDataCells(sheet, headerRow + 1, endRow, run.StartColumn, run.EndColumn) + run.Count * 3)
                .FirstOrDefault();

            if (bestRun == null)
            {
                return range;
            }

            var dataStartRow = FindFirstDataRow(sheet, headerRow + 1, endRow, bestRun.StartColumn, bestRun.EndColumn);
            if (dataStartRow <= 0)
            {
                return range;
            }

            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = dataStartRow,
                EndRow = range.EndRow,
                StartColumn = bestRun.StartColumn,
                EndColumn = bestRun.EndColumn
            };
        }

        private static CellRange FindAverageRange(SheetSnapshot sheet, int startRow, int endRow, CellRange measurementRange)
        {
            var range = FindDataRange(sheet, startRow, endRow, AverageKeywords);
            if (range != null)
            {
                return range;
            }

            if (measurementRange == null)
            {
                return null;
            }

            var headerRow = FindAverageSubHeaderRow(
                sheet,
                startRow,
                Math.Max(startRow, measurementRange.StartRow - 1),
                measurementRange.StartColumn,
                measurementRange.EndColumn);
            if (headerRow <= 0)
            {
                return null;
            }

            var columns = sheet.Cells
                .Where(cell =>
                    cell.Row == headerRow &&
                    cell.Column >= measurementRange.StartColumn &&
                    cell.Column <= measurementRange.EndColumn &&
                    IsAverageHeader(cell.Text))
                .OrderBy(cell => cell.Column)
                .Select(cell => cell.Column)
                .Distinct()
                .ToList();

            if (columns.Count == 0)
            {
                return null;
            }

            var dataStartRow = FindFirstDataRow(sheet, headerRow + 1, endRow, columns.First(), columns.Last());
            if (dataStartRow <= 0)
            {
                dataStartRow = measurementRange.StartRow;
            }

            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = dataStartRow,
                EndRow = measurementRange.EndRow,
                StartColumn = columns.First(),
                EndColumn = columns.Last()
            };
        }

        private static int FindAverageSubHeaderRow(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            for (var row = startRow; row <= endRow; row++)
            {
                var hasAverageHeader = sheet.Cells.Any(cell =>
                    cell.Row == row &&
                    cell.Column >= startColumn &&
                    cell.Column <= endColumn &&
                    IsAverageHeader(cell.Text));
                if (hasAverageHeader)
                {
                    return row;
                }
            }

            return 0;
        }

        private static bool IsAverageHeader(string text)
        {
            return AverageKeywords.Any(keyword => MatchesKeyword(text, keyword));
        }

        private static CellRange ExcludeAverageColumns(CellRange measurementRange, CellRange averageRange)
        {
            if (measurementRange == null || averageRange == null)
            {
                return measurementRange;
            }

            if (measurementRange.EndColumn < averageRange.StartColumn ||
                measurementRange.StartColumn > averageRange.EndColumn)
            {
                return measurementRange;
            }

            if (averageRange.StartColumn <= measurementRange.StartColumn &&
                averageRange.EndColumn >= measurementRange.EndColumn)
            {
                return null;
            }

            var leftWidth = Math.Max(0, averageRange.StartColumn - measurementRange.StartColumn);
            var rightWidth = Math.Max(0, measurementRange.EndColumn - averageRange.EndColumn);
            var keepLeftSide = leftWidth >= rightWidth;

            return new CellRange
            {
                SheetName = measurementRange.SheetName,
                StartRow = measurementRange.StartRow,
                EndRow = measurementRange.EndRow,
                StartColumn = keepLeftSide ? measurementRange.StartColumn : averageRange.EndColumn + 1,
                EndColumn = keepLeftSide ? averageRange.StartColumn - 1 : measurementRange.EndColumn
            };
        }

    }
}
