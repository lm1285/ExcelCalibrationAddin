using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelCalibrationAddin.Host.Services
{
    public static class ManualStandardValueRangeParser
    {
        private const string NumberPattern = @"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?";
        private static readonly Regex RangeRegex = new Regex(
            @"^\s*(?<lower>" + NumberPattern + @")\s*(?:~|～|至|到|—|–|(?<![eE])-|－)\s*(?<upper>" + NumberPattern + @")\s*$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static bool TryParse(string text, out double lowerBound, out double upperBound)
        {
            lowerBound = 0;
            upperBound = 0;
            var match = RangeRegex.Match(text ?? string.Empty);
            if (!match.Success ||
                !TryParseNumber(match.Groups["lower"].Value, out var first) ||
                !TryParseNumber(match.Groups["upper"].Value, out var second) ||
                double.IsNaN(first) ||
                double.IsInfinity(first) ||
                double.IsNaN(second) ||
                double.IsInfinity(second))
            {
                return false;
            }

            lowerBound = Math.Min(first, second);
            upperBound = Math.Max(first, second);
            return true;
        }

        private static bool TryParseNumber(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
                double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }
    }
}
