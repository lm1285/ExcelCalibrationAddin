using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class ParameterValueParser
    {
        private static readonly Regex NumberRegex = new Regex(@"[-+]?\d+(\.\d+)?([eE][-+]?\d+)?", RegexOptions.Compiled);

        public double ParseNumeric(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                throw new InvalidOperationException("Parameter value is empty.");
            }

            var normalized = Normalize(raw);
            var match = NumberRegex.Match(normalized);
            if (!match.Success)
            {
                throw new InvalidOperationException($"Cannot parse numeric value from '{raw}'.");
            }

            return double.Parse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture);
        }

        public double ParseMpe(string raw, bool percentageAsRatio)
        {
            var value = ParseNumeric(raw);
            if (percentageAsRatio && Normalize(raw).Contains("%"))
            {
                return value / 100.0;
            }

            return value;
        }

        private static string Normalize(string raw)
        {
            return (raw ?? string.Empty)
                .Trim()
                .Replace("％", "%")
                .Replace("＋", "+")
                .Replace("－", "-");
        }
    }
}
