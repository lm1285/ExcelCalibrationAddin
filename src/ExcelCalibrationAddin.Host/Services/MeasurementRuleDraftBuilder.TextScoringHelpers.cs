using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleDraftBuilder
    {
        private static double ScoreKeywordMatch(string text, string keyword)
        {
            if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(keyword))
            {
                return 0;
            }

            var value = NormalizeHeaderText(text);
            var normalizedKeyword = NormalizeHeaderText(keyword);
            if (string.Equals(value, normalizedKeyword, StringComparison.OrdinalIgnoreCase))
            {
                return 20;
            }

            return value.IndexOf(normalizedKeyword, StringComparison.OrdinalIgnoreCase) >= 0 ? 12 : 0;
        }

        private static bool MatchesKeyword(string text, string keyword)
        {
            return ScoreKeywordMatch(text, keyword) > 0;
        }

        private static int InferPrintEndColumn(SheetSnapshot sheet)
        {
            if (sheet.Cells.Count == 0)
            {
                return 1;
            }

            return sheet.Cells.Max(cell => cell.MergeRange?.EndColumn ?? cell.Column);
        }

        private static bool LooksLikeWrongFieldHeader(string text, string[] keywords)
        {
            var value = NormalizeHeaderText(text);
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var isMeasurementSearch = keywords == MeasurementKeywords;
            if (isMeasurementSearch &&
                (value.IndexOf("\u8BEF\u5DEE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 value.IndexOf("\u5E73\u5747", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 value.IndexOf("AVG", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 value.IndexOf("\u6280\u672F\u8981\u6C42", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 value.IndexOf("\u7ED3\u8BBA", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 value.IndexOf("P/F", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return true;
            }

            if (isMeasurementSearch &&
                (value.IndexOf("\u6807\u51C6\u5668", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 value.IndexOf("\u53C2\u8003\u5668", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 value.IndexOf("\u6807\u51C6\u6D4B\u91CF", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return true;
            }

            return false;
        }

        private static string NormalizeHeaderText(string text)
        {
            var value = (text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var chars = value
                .Where(ch => !char.IsWhiteSpace(ch) &&
                             ch != '(' && ch != ')' &&
                             ch != '\uFF08' && ch != '\uFF09' &&
                             ch != ':' && ch != '\uFF1A' &&
                             ch != '/' && ch != '\\')
                .ToArray();
            return new string(chars);
        }

        private static int ScoreAnchorDistance(CellMeta cell, params CellRange[] anchorRanges)
        {
            var anchors = anchorRanges?.Where(range => range != null).ToList();
            if (anchors == null || anchors.Count == 0)
            {
                return 0;
            }

            return anchors.Min(range => Math.Abs((range.StartColumn + range.EndColumn) / 2 - cell.Column));
        }

        private static bool LooksLikeSectionTitle(string text)
        {
            var value = (text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (value.Contains('\uFF1A') || value.Contains(':'))
            {
                var prefix = value.Split(new[] { '\uFF1A', ':' }, 2)[0];
                if (prefix.Any(char.IsDigit) || StartsWithChineseSectionNumber(prefix))
                {
                    return true;
                }
            }

            return value.Length <= 16 && StartsWithChineseSectionNumber(value);
        }

        private static bool LooksLikeInlineNote(string text)
        {
            var value = (text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if ((value.Contains('\uFF1A') || value.Contains(':')) &&
                (value.Contains("~") || value.Contains('\uFF5E') || value.Contains('(') || value.Contains(')') || value.Contains('\uFF08') || value.Contains('\uFF09')))
            {
                return true;
            }

            return value.Any(char.IsDigit) && (value.Contains("~") || value.Contains('\uFF5E'));
        }

        private static bool StartsWithChineseSectionNumber(string text)
        {
            return !string.IsNullOrWhiteSpace(text) && "\u4E00\u4E8C\u4E09\u56DB\u4E94\u516D\u4E03\u516B\u4E5D\u5341".IndexOf(text[0]) >= 0;
        }

        private static int FindSubHeaderRow(SheetSnapshot sheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            var lastRow = Math.Min(endRow, startRow + 2);
            for (var row = startRow; row <= lastRow; row++)
            {
                var headerCells = sheet.Cells
                    .Where(cell =>
                        cell.Row == row &&
                        cell.Column >= startColumn &&
                        cell.Column <= endColumn &&
                        !string.IsNullOrWhiteSpace(cell.Text))
                    .OrderBy(cell => cell.Column)
                    .ToList();

                if (headerCells.Count(cell => IsMeasurementSubHeader(cell.Text)) >= 1)
                {
                    return row;
                }
            }

            return 0;
        }

        private static bool IsMeasurementSubHeader(string text)
        {
            var value = (text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (MeasurementSubHeaderExcludes.Any(item => value.IndexOf(item, StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return false;
            }

            if (int.TryParse(value, out _))
            {
                return true;
            }

            if (MeasurementAttemptHeaderRegex.IsMatch(NormalizeHeaderText(value)))
            {
                return true;
            }

            return value.Length <= 4 && value.All(char.IsDigit);
        }

        private static string BuildNotes(SheetSnapshot sheet, int startRow, int endRow)
        {
            var texts = sheet.Cells
                .Where(cell =>
                    cell.Row >= startRow &&
                    cell.Row <= Math.Min(endRow, startRow + 2) &&
                    cell.Column <= 8 &&
                    !string.IsNullOrWhiteSpace(cell.Text))
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .Select(cell => cell.Text.Trim())
                .Distinct()
                .Take(4)
                .ToList();

            return string.Join(" / ", texts);
        }

    }
}
