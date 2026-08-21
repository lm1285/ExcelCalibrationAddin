using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Recognition
{
    public static class RequirementTextParser
    {
        private static readonly Regex NumberRegex = new Regex(@"[-+]?\d+(\.\d+)?([eE][-+]?\d+)?", RegexOptions.Compiled);

        public static ParsedTechnicalRequirement Parse(CellMeta cell)
        {
            var displayText = ResolvePreferredText(cell);
            var rawValueText = cell?.RawValueText ?? string.Empty;
            var numberFormat = cell?.NumberFormat ?? string.Empty;
            var formatSignals = ExtractNumberFormatSignals(numberFormat);

            var textBuilder = new StringBuilder();
            AddRequirementText(textBuilder, displayText);

            if (!ContainsNumber(displayText))
            {
                AddRequirementText(textBuilder, rawValueText);
            }

            var text = textBuilder.ToString();
            if (string.IsNullOrWhiteSpace(text))
            {
                text = rawValueText;
            }

            var operatorContext = string.Join(" ", new[] { displayText, rawValueText, formatSignals }
                .Where(item => !string.IsNullOrWhiteSpace(item)));

            return new ParsedTechnicalRequirement
            {
                Text = text,
                Operator = ResolveRequirementOperator(operatorContext),
                ContextSignals = new List<string>
                {
                    displayText,
                    rawValueText,
                    formatSignals,
                    numberFormat
                }
            };
        }

        private static string ResolvePreferredText(CellMeta cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            if (!string.IsNullOrWhiteSpace(cell.DisplayText) &&
                !LooksLikeUnavailableDisplayText(cell.DisplayText))
            {
                return cell.DisplayText.Trim();
            }

            if (!string.IsNullOrWhiteSpace(cell.Text) &&
                !LooksLikeUnavailableDisplayText(cell.Text))
            {
                return cell.Text.Trim();
            }

            return cell.RawValueText ?? string.Empty;
        }

        private static bool LooksLikeUnavailableDisplayText(string text)
        {
            var value = (text ?? string.Empty).Trim();
            return value.Length > 0 && value.All(ch => ch == '#');
        }

        private static void AddRequirementText(StringBuilder builder, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            if (builder.Length > 0)
            {
                builder.Append(' ');
            }

            builder.Append(text.Trim());
        }

        private static string ExtractNumberFormatSignals(string numberFormat)
        {
            if (string.IsNullOrWhiteSpace(numberFormat))
            {
                return string.Empty;
            }

            var section = numberFormat.Split(';')[0];
            var signals = new List<string>();
            var inQuote = false;
            var buffer = new StringBuilder();

            foreach (var ch in section)
            {
                if (ch == '"')
                {
                    if (inQuote && buffer.Length > 0)
                    {
                        signals.Add(buffer.ToString());
                    }

                    buffer.Clear();
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote)
                {
                    buffer.Append(ch);
                    continue;
                }

                if (IsRequirementSignal(ch))
                {
                    signals.Add(ch.ToString());
                }
            }

            return string.Join(string.Empty, signals).Trim();
        }

        private static bool IsRequirementSignal(char ch)
        {
            return ch == '%' ||
                ch == '<' ||
                ch == '>' ||
                ch == '=' ||
                ch == '+' ||
                ch == '-' ||
                ch == '\u00B1' ||
                ch == '\u2264' ||
                ch == '\u2265' ||
                ch == '\uFF1C' ||
                ch == '\uFF1E' ||
                ch == '\uFF05' ||
                ch == '\uFF0B' ||
                ch == '\uFF0D';
        }

        private static TechnicalRequirementOperator ResolveRequirementOperator(string text)
        {
            var normalized = NormalizeSignal(text)
                .Replace("\uFF1C", "<")
                .Replace("\uFF1E", ">")
                .Replace("\u2264", "<=")
                .Replace("\u2265", ">=");

            if (normalized.Contains("+/-"))
            {
                return TechnicalRequirementOperator.PlusMinus;
            }

            if (normalized.Contains("<=") || normalized.Contains("=<"))
            {
                return TechnicalRequirementOperator.LessThanOrEqual;
            }

            if (normalized.Contains(">=") || normalized.Contains("=>"))
            {
                return TechnicalRequirementOperator.GreaterThanOrEqual;
            }

            if (normalized.Contains("<"))
            {
                return TechnicalRequirementOperator.LessThan;
            }

            if (normalized.Contains(">"))
            {
                return TechnicalRequirementOperator.GreaterThan;
            }

            return TechnicalRequirementOperator.None;
        }

        private static bool ContainsNumber(string text)
        {
            return !string.IsNullOrWhiteSpace(text) && NumberRegex.IsMatch(text);
        }

        private static string NormalizeSignal(string text)
        {
            return (text ?? string.Empty)
                .Trim()
                .Replace("\uFF05", "%")
                .Replace("\uFF08", "(")
                .Replace("\uFF09", ")")
                .Replace("\uFF0D", "-")
                .Replace("\u2014", "-")
                .Replace("\u00B1", "+/-")
                .ToUpperInvariant();
        }
    }

    public sealed class ParsedTechnicalRequirement
    {
        public string Text { get; set; } = string.Empty;
        public TechnicalRequirementOperator Operator { get; set; }
        public List<string> ContextSignals { get; set; } = new List<string>();
    }
}
