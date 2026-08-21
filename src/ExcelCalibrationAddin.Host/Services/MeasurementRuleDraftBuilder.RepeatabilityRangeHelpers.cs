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
        private static CellRange RefineRepeatabilityMeasurementRange(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            CellRange measurementRange,
            CellRange standardRange,
            CellRange averageRange)
        {
            var candidate = FindRepeatabilityMeasurementByHeader(sheet, startRow, endRow, standardRange, averageRange)
                ?? measurementRange;

            if (candidate != null)
            {
                var row = FindBestRepeatabilityDataRow(sheet, candidate, endRow);
                if (row > 0)
                {
                    var run = FindNumericRunInRow(sheet, row, candidate.StartColumn, candidate.EndColumn, standardRange, averageRange);
                    if (run != null)
                    {
                        return new CellRange
                        {
                            SheetName = sheet.Name,
                            StartRow = row,
                            EndRow = row,
                            StartColumn = run.StartColumn,
                            EndColumn = run.EndColumn
                        };
                    }

                    return new CellRange
                    {
                        SheetName = candidate.SheetName,
                        StartRow = row,
                        EndRow = row,
                        StartColumn = candidate.StartColumn,
                        EndColumn = candidate.EndColumn
                    };
                }
            }

            return InferRepeatabilityMeasurementWithoutHeader(sheet, startRow, endRow, standardRange, averageRange);
        }

        private static CellRange FindRepeatabilityMeasurementByHeader(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            CellRange standardRange,
            CellRange averageRange)
        {
            var headerCandidates = sheet.Cells
                .Where(cell =>
                    cell.Row >= startRow &&
                    cell.Row <= Math.Min(endRow, startRow + 8) &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    IsRepeatabilityMeasurementHeader(cell.Text))
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();

            foreach (var header in headerCandidates)
            {
                var headerStartColumn = header.MergeRange?.StartColumn ?? header.Column;
                var headerEndColumn = header.MergeRange?.EndColumn ?? header.Column;
                var headerBottomRow = header.MergeRange?.EndRow ?? header.Row;

                if (headerStartColumn == headerEndColumn)
                {
                    headerEndColumn = Math.Max(headerEndColumn, FindRightBoundaryBeforeKnownFields(sheet, header.Row, headerEndColumn, endRow, averageRange));
                }

                var dataStartRow = FindFirstRepeatabilityDataRow(sheet, headerBottomRow + 1, endRow, headerStartColumn, headerEndColumn);
                if (dataStartRow <= 0)
                {
                    continue;
                }

                var run = FindNumericRunInRow(sheet, dataStartRow, headerStartColumn, headerEndColumn, standardRange, averageRange);
                var startColumn = run?.StartColumn ?? headerStartColumn;
                var endColumn = run?.EndColumn ?? headerEndColumn;

                return new CellRange
                {
                    SheetName = sheet.Name,
                    StartRow = dataStartRow,
                    EndRow = dataStartRow,
                    StartColumn = startColumn,
                    EndColumn = endColumn
                };
            }

            return null;
        }

        private static CellRange InferRepeatabilityMeasurementWithoutHeader(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            CellRange standardRange,
            CellRange averageRange)
        {
            var searchStartRow = Math.Max(startRow + 1, Math.Min(
                standardRange?.StartRow ?? int.MaxValue,
                averageRange?.StartRow ?? int.MaxValue));
            if (searchStartRow == int.MaxValue)
            {
                searchStartRow = startRow + 1;
            }

            NumericHeaderRun best = null;
            for (var row = searchStartRow; row <= endRow; row++)
            {
                var run = FindBestNumericRunInRow(sheet, row, 1, InferMaxColumn(sheet), standardRange, averageRange);
                if (run == null || run.Count < 2)
                {
                    continue;
                }

                var score = (run.EndColumn - run.StartColumn + 1) * 10 - Math.Max(0, row - searchStartRow);
                if (best == null || score > best.Score)
                {
                    best = new NumericHeaderRun
                    {
                        HeaderRow = row - 1,
                        StartColumn = run.StartColumn,
                        EndColumn = run.EndColumn,
                        DataStartRow = row,
                        Score = score
                    };
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
                EndRow = best.DataStartRow,
                StartColumn = best.StartColumn,
                EndColumn = best.EndColumn
            };
        }

    }
}
