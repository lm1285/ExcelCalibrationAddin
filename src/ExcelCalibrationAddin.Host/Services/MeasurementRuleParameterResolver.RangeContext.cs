using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleParameterResolver
    {
        private static readonly Regex RangeIntervalRegex = new Regex(
            @"(?<start>[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)\s*[^\d+\-~\uFF5E\u81F3\u2013\u2014]*?\s*(?:~|\uFF5E|\u81F3|\u2013|\u2014|-)\s*[^\d+\-]*?\s*(?<end>[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)",
            RegexOptions.Compiled);

        private static double? ResolveReferenceRange(SheetSnapshot sheet, CellRange range)
        {
            if (sheet == null || range == null)
            {
                return null;
            }

            var logicalCells = MergedCellLogicalRangeResolver.GetTextCells(sheet, range);
            if (logicalCells.Count == 0)
            {
                return null;
            }

            var combinedBuilder = new StringBuilder();
            foreach (var logicalCell in logicalCells.OrderBy(item => item.Range.StartRow).ThenBy(item => item.Range.StartColumn))
            {
                if (!string.IsNullOrWhiteSpace(logicalCell.Anchor?.Text))
                {
                    if (combinedBuilder.Length > 0)
                    {
                        combinedBuilder.Append(' ');
                    }

                    combinedBuilder.Append(logicalCell.Anchor.Text.Trim());
                }
            }

            var combinedText = combinedBuilder.ToString();
            if (string.IsNullOrWhiteSpace(combinedText))
            {
                return null;
            }

            var intervalMatch = RangeIntervalRegex.Match(combinedText);
            if (intervalMatch.Success &&
                double.TryParse(intervalMatch.Groups["start"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var start) &&
                double.TryParse(intervalMatch.Groups["end"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var end))
            {
                var span = ResolveIntervalReferenceRange(start, end);
                if (span > 0)
                {
                    return span;
                }
            }

            var matches = NumberRegex.Matches(combinedText)
                .Cast<Match>()
                .Select(match => double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                    ? (double?)value
                    : null)
                .Where(value => value.HasValue)
                .Select(value => value.Value)
                .ToList();

            if (matches.Count == 0)
            {
                return null;
            }

            return matches.Max(value => Math.Abs(value));
        }

        private static double ResolveIntervalReferenceRange(double start, double end)
        {
            var span = Math.Abs(end - start);
            if (span <= 0)
            {
                return 0;
            }

            var startMagnitude = Math.Abs(start);
            var endMagnitude = Math.Abs(end);
            var maxMagnitude = Math.Max(startMagnitude, endMagnitude);
            if (start < 0 &&
                end > 0 &&
                maxMagnitude > 0 &&
                Math.Abs(startMagnitude - endMagnitude) <= Math.Max(1e-9, maxMagnitude * 1e-9))
            {
                return maxMagnitude;
            }

            return span;
        }

        private static List<string> CollectRangeContextTexts(SheetSnapshot sheet, CellRange range)
        {
            var texts = new List<string>();
            if (sheet == null || range == null)
            {
                return texts;
            }

            var headerStartRow = Math.Max(1, range.StartRow - 4);
            foreach (var cell in MergedCellLogicalRangeResolver.GetTextCells(sheet, new CellRange
            {
                SheetName = range.SheetName,
                StartRow = headerStartRow,
                EndRow = Math.Max(headerStartRow, range.StartRow - 1),
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            }))
            {
                AddContextText(texts, cell.Anchor?.Text);
                AddContextText(texts, cell.Anchor?.NumberFormat);
            }

            return texts;
        }

        private static List<string> CollectCellContextTexts(SheetSnapshot sheet, LogicalCellRange logicalCell)
        {
            var texts = new List<string>();
            if (sheet == null || logicalCell == null)
            {
                return texts;
            }

            AddContextText(texts, logicalCell.Anchor?.Text);
            AddContextText(texts, logicalCell.Anchor?.NumberFormat);

            var rowStart = logicalCell.Range.StartRow;
            var rowEnd = logicalCell.Range.EndRow;
            var leftBoundary = Math.Max(1, logicalCell.Range.StartColumn - 4);
            foreach (var neighbor in sheet.Cells
                .Where(cell =>
                    cell.Row >= rowStart &&
                    cell.Row <= rowEnd &&
                    cell.Column >= leftBoundary &&
                    cell.Column < logicalCell.Range.StartColumn &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    LooksLikeLabelContext(cell.Text))
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column))
            {
                AddContextText(texts, neighbor.Text);
            }

            return texts;
        }

        private static void AddContextText(ICollection<string> texts, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            var trimmed = text.Trim();
            if (trimmed.Length == 0 || texts.Contains(trimmed))
            {
                return;
            }

            texts.Add(trimmed);
        }

        private static bool ContainsPercentToken(string text)
        {
            return !string.IsNullOrWhiteSpace(text) && NormalizeSignal(text).Contains("%");
        }

        private static bool LooksLikeLabelContext(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var trimmed = text.Trim();
            if (ReferencedMarkers.Any(marker =>
                    string.Equals(NormalizeSignal(trimmed), NormalizeSignal(marker), StringComparison.Ordinal)))
            {
                return false;
            }

            var digitCount = trimmed.Count(char.IsDigit);
            if (digitCount > 0)
            {
                var normalized = NormalizeSignal(trimmed);
                return normalized.Contains("技术要求") ||
                    normalized.Contains("允许误差") ||
                    normalized.Contains("允差") ||
                    normalized.Contains("量程") ||
                    normalized.Contains("MPE");
            }

            return true;
        }
        private static string NormalizeSignal(string text)
        {
            return (text ?? string.Empty)
                .Trim()
                .Replace("\uFF05", "%")
                .Replace("\uFF08", "(")
                .Replace("\uFF09", ")")
                .Replace("\uFF0D", "-")
                .Replace("\u2013", "-")
                .Replace("\u2014", "-")
                .Replace("\u00B1", "+/-")
                .ToUpperInvariant();
        }
    }
}
