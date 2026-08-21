using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelCalibrationAddin.Host.Services
{
    internal static class TemplateUnitParser
    {
        private static readonly Regex UnitRegex = new Regex(
            @"%\s*FS|FS\s*%|%\s*LEL|%\s*VOL|PPM|PPB|[UµμΜ]?MOL\s*/\s*MOL|MG\s*/\s*M(?:3|³)|[UµμΜ]G\s*/\s*M(?:3|³)|MPA|KPA|PA|°C|℃|秒|(?<![A-Z])MS(?![A-Z])|(?<![A-Z])S(?![A-Z])|%",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static string Extract(params string[] candidates)
        {
            return Extract((IEnumerable<string>)candidates);
        }

        public static string Extract(IEnumerable<string> candidates)
        {
            foreach (var candidate in candidates ?? Enumerable.Empty<string>())
            {
                var normalized = Normalize(candidate);
                var match = UnitRegex.Match(normalized);
                if (match.Success)
                {
                    return Canonicalize(match.Value);
                }

                var parenthesized = Regex.Match(candidate ?? string.Empty, @"[\(（](?<unit>[^\)）]{1,20})[\)）]");
                if (parenthesized.Success && LooksLikeGenericUnit(parenthesized.Groups["unit"].Value))
                {
                    return parenthesized.Groups["unit"].Value.Trim();
                }

                var trimmed = (candidate ?? string.Empty).Trim().Trim('"');
                if (trimmed.Length <= 12 && LooksLikeGenericUnit(trimmed))
                {
                    return trimmed;
                }
            }

            return string.Empty;
        }

        private static bool LooksLikeGenericUnit(string value)
        {
            var text = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text) ||
                Regex.IsMatch(text, @"^[-+]?\d+(?:\.\d+)?$"))
            {
                return false;
            }

            return Regex.IsMatch(text, @"[A-Za-zµμΜ°%/]") &&
                Regex.IsMatch(text, @"^[A-Za-z0-9µμΜ°%/²³.\-\s]+$");
        }

        public static bool SameUnitFamily(string left, string right)
        {
            var normalizedLeft = Canonicalize(left);
            var normalizedRight = Canonicalize(right);
            return string.IsNullOrWhiteSpace(normalizedLeft) ||
                string.IsNullOrWhiteSpace(normalizedRight) ||
                string.Equals(normalizedLeft, normalizedRight, StringComparison.OrdinalIgnoreCase);
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty)
                .Replace("％", "%")
                .Replace("／", "/")
                .Replace("㎥", "m3")
                .ToUpperInvariant();
        }

        private static string Canonicalize(string value)
        {
            var normalized = Normalize(value).Replace(" ", string.Empty);
            if (string.IsNullOrWhiteSpace(normalized)) return string.Empty;
            if (normalized == "%FS" || normalized == "FS%") return "%FS";
            if (normalized == "%LEL") return "%LEL";
            if (normalized == "%VOL") return "%VOL";
            if (normalized == "%") return "%";
            if (normalized == "PPM") return "ppm";
            if (normalized == "PPB") return "ppb";
            if (normalized.EndsWith("MOL/MOL", StringComparison.Ordinal))
            {
                return normalized == "MOL/MOL" ? "mol/mol" : "μmol/mol";
            }
            if (normalized == "MG/M3" || normalized == "MG/M³") return "mg/m3";
            if (normalized == "UG/M3" || normalized == "UG/M³" ||
                normalized == "µG/M3" || normalized == "μG/M3" ||
                normalized == "ΜG/M3") return "μg/m3";
            if (normalized == "MPA") return "MPa";
            if (normalized == "KPA") return "kPa";
            if (normalized == "PA") return "Pa";
            if (normalized == "°C" || normalized == "℃") return "°C";
            if (normalized == "MS") return "ms";
            if (normalized == "S" || normalized == "秒") return "s";
            return value?.Trim() ?? string.Empty;
        }
    }
}
