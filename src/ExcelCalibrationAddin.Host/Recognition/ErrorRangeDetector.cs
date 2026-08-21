using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;

namespace ExcelCalibrationAddin.Host.Recognition
{
    public sealed class ErrorRangeDetector
    {
        private readonly NumberFormatInterpreter _numberFormatInterpreter;

        public ErrorRangeDetector(NumberFormatInterpreter numberFormatInterpreter)
        {
            _numberFormatInterpreter = numberFormatInterpreter;
        }

        public CellRange Infer(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            CellRange measurementRange,
            CellRange averageRange,
            CellRange standardRange)
        {
            var referenceRange = averageRange ?? measurementRange ?? standardRange;
            if (referenceRange == null)
            {
                return null;
            }

            var referenceCells = GetRangeCells(sheet, referenceRange)
                .Where(cell => !string.IsNullOrWhiteSpace(cell.NumberFormat) || !string.IsNullOrWhiteSpace(cell.Text))
                .ToList();
            var referenceFormat = ResolveDominantFormat(referenceCells.Select(cell => cell.NumberFormat));
            var referenceRule = _numberFormatInterpreter.Interpret(referenceFormat);
            var dataStartRow = referenceRange.StartRow;
            var dataEndRow = endRow;
            var searchStartColumn = Math.Max(1, referenceRange.EndColumn - 1);
            var searchEndColumn = Math.Min(InferMaxColumn(sheet), referenceRange.EndColumn + 6);

            ColumnScore best = null;
            for (var column = searchStartColumn; column <= searchEndColumn; column++)
            {
                if (IsColumnInsideRange(column, referenceRange) ||
                    IsColumnInsideRange(column, averageRange) ||
                    IsColumnInsideRange(column, standardRange))
                {
                    continue;
                }

                var score = ScoreColumn(sheet, column, dataStartRow, dataEndRow, referenceRange, referenceRule, referenceFormat);
                if (score != null && (best == null || score.Score > best.Score))
                {
                    best = score;
                }
            }

            if (best == null || best.Score < 45)
            {
                return null;
            }

            var startColumn = best.Column;
            var endColumn = best.Column;
            for (var column = best.Column + 1; column <= searchEndColumn; column++)
            {
                var sibling = ScoreColumn(sheet, column, dataStartRow, dataEndRow, referenceRange, referenceRule, referenceFormat);
                if (!IsSiblingColumn(best, sibling))
                {
                    break;
                }

                endColumn = column;
            }

            var actualStartRow = FindFirstDataRow(sheet, dataStartRow, dataEndRow, startColumn, endColumn);
            if (actualStartRow <= 0)
            {
                actualStartRow = dataStartRow;
            }

            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = actualStartRow,
                EndRow = dataEndRow,
                StartColumn = startColumn,
                EndColumn = endColumn
            };
        }

        private ColumnScore ScoreColumn(
            SheetSnapshot sheet,
            int column,
            int startRow,
            int endRow,
            CellRange referenceRange,
            FormatRule referenceRule,
            string referenceFormat)
        {
            var cells = sheet.Cells
                .Where(cell => cell.Column == column && cell.Row >= startRow && cell.Row <= endRow)
                .OrderBy(cell => cell.Row)
                .ToList();
            if (cells.Count == 0)
            {
                return null;
            }

            var nonEmpty = cells
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Text) || !string.IsNullOrWhiteSpace(cell.Formula))
                .ToList();
            if (nonEmpty.Count == 0)
            {
                return null;
            }

            var numericCount = nonEmpty.Count(cell => SheetRowContentAnalyzer.LooksNumeric(cell.Text));
            var formulaCount = nonEmpty.Count(cell => !string.IsNullOrWhiteSpace(cell.Formula));
            var format = ResolveDominantFormat(nonEmpty.Select(cell => cell.NumberFormat));
            var rule = _numberFormatInterpreter.Interpret(format);
            var score = 0d;

            score += Math.Min(24, numericCount * 4);
            score += Math.Min(20, formulaCount * 6);
            score += Math.Max(0, 18 - Math.Abs(column - referenceRange.EndColumn) * 4);

            if (!string.IsNullOrWhiteSpace(format))
            {
                score += 8;
            }

            if (!string.IsNullOrWhiteSpace(referenceFormat) &&
                string.Equals(NormalizeFormat(format), NormalizeFormat(referenceFormat), StringComparison.OrdinalIgnoreCase))
            {
                score += 18;
            }

            if (rule.DecimalPlaces == referenceRule.DecimalPlaces)
            {
                score += 10;
            }

            if (rule.IsPercentage)
            {
                score += 8;
            }

            if (nonEmpty.Count >= Math.Max(2, (endRow - startRow + 1) / 3))
            {
                score += 6;
            }

            return new ColumnScore
            {
                Column = column,
                Score = score,
                DecimalPlaces = rule.DecimalPlaces,
                IsPercentage = rule.IsPercentage,
                NumberFormat = format
            };
        }

        private static bool IsSiblingColumn(ColumnScore seed, ColumnScore candidate)
        {
            if (seed == null || candidate == null || candidate.Score < 35)
            {
                return false;
            }

            return seed.IsPercentage == candidate.IsPercentage &&
                seed.DecimalPlaces == candidate.DecimalPlaces &&
                string.Equals(
                    NormalizeFormat(seed.NumberFormat),
                    NormalizeFormat(candidate.NumberFormat),
                    StringComparison.OrdinalIgnoreCase);
        }

        private static int FindFirstDataRow(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            for (var row = startRow; row <= endRow; row++)
            {
                if (sheet.Cells.Any(cell =>
                    cell.Row == row &&
                    cell.Column >= startColumn &&
                    cell.Column <= endColumn &&
                    SheetRowContentAnalyzer.CellHasEffectiveContent(sheet, cell)))
                {
                    return row;
                }
            }

            return 0;
        }

        private static IEnumerable<CellMeta> GetRangeCells(SheetSnapshot sheet, CellRange range)
        {
            if (range == null)
            {
                return Enumerable.Empty<CellMeta>();
            }

            return sheet.Cells.Where(cell =>
                cell.Row >= range.StartRow &&
                cell.Row <= range.EndRow &&
                cell.Column >= range.StartColumn &&
                cell.Column <= range.EndColumn);
        }

        private static int InferMaxColumn(SheetSnapshot sheet)
        {
            return sheet.Cells.Count == 0 ? 1 : sheet.Cells.Max(item => item.Column);
        }

        private static bool IsColumnInsideRange(int column, CellRange range)
        {
            return range != null && column >= range.StartColumn && column <= range.EndColumn;
        }

        private static string ResolveDominantFormat(IEnumerable<string> formats)
        {
            return formats
                .Where(format => !string.IsNullOrWhiteSpace(format))
                .GroupBy(NormalizeFormat)
                .OrderByDescending(group => group.Count())
                .Select(group => group.First())
                .FirstOrDefault() ?? string.Empty;
        }

        private static string NormalizeFormat(string format)
        {
            return (format ?? string.Empty).Trim().Replace("_", string.Empty);
        }

        private sealed class ColumnScore
        {
            public int Column { get; set; }
            public double Score { get; set; }
            public string NumberFormat { get; set; } = string.Empty;
            public int? DecimalPlaces { get; set; }
            public bool IsPercentage { get; set; }
        }
    }
}
