using System;
using System.Globalization;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    internal static class MpeValuePatternCodec
    {
        private const string Prefix = "mpe:";

        public static string Build(
            ErrorType errorType,
            double scaleFactor,
            TechnicalRequirementOperator requirementOperator = TechnicalRequirementOperator.None,
            string unit = null)
        {
            var pattern = Prefix +
                errorType.ToString().ToLowerInvariant() +
                ":scale=" +
                scaleFactor.ToString("G17", CultureInfo.InvariantCulture);

            if (requirementOperator != TechnicalRequirementOperator.None)
            {
                pattern += ":op=" + requirementOperator.ToString().ToLowerInvariant();
            }

            var normalizedUnit = NormalizeUnit(unit);
            if (!string.IsNullOrWhiteSpace(normalizedUnit))
            {
                pattern += ":unit=" + normalizedUnit;
            }

            return pattern;
        }

        public static MpeValuePattern Parse(string valuePattern)
        {
            if (string.IsNullOrWhiteSpace(valuePattern) ||
                !valuePattern.Trim().StartsWith(Prefix, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var parts = valuePattern.Trim().Split(':');
            if (parts.Length < 3 ||
                !Enum.TryParse(parts[1], true, out ErrorType errorType))
            {
                return null;
            }

            var scaleFactor = 1d;
            foreach (var part in parts.Skip(2))
            {
                if (!part.StartsWith("scale=", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var rawScale = part.Substring("scale=".Length);
                if (!double.TryParse(rawScale, NumberStyles.Float, CultureInfo.InvariantCulture, out scaleFactor) ||
                    double.IsNaN(scaleFactor) ||
                    double.IsInfinity(scaleFactor) ||
                    scaleFactor <= 0)
                {
                    return null;
                }
            }

            var requirementOperator = TechnicalRequirementOperator.None;
            var unit = string.Empty;
            foreach (var part in parts.Skip(2))
            {
                if (!part.StartsWith("op=", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Enum.TryParse(part.Substring("op=".Length), true, out requirementOperator);
            }

            foreach (var part in parts.Skip(2))
            {
                if (part.StartsWith("unit=", StringComparison.OrdinalIgnoreCase))
                {
                    unit = NormalizeUnit(part.Substring("unit=".Length));
                }
            }

            return new MpeValuePattern
            {
                ErrorType = errorType,
                ScaleFactor = scaleFactor,
                Operator = requirementOperator,
                Unit = unit,
                RawPattern = valuePattern.Trim()
            };
        }

        internal static string NormalizeUnit(string value)
        {
            var normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
            return normalized == "s" || normalized == "sec" || normalized == "second" ||
                normalized == "seconds" || normalized == "秒"
                ? "s"
                : string.Empty;
        }
    }

    internal sealed class MpeValuePattern
    {
        public ErrorType ErrorType { get; set; }
        public double ScaleFactor { get; set; } = 1d;
        public TechnicalRequirementOperator Operator { get; set; }
        public string Unit { get; set; } = string.Empty;
        public string RawPattern { get; set; } = string.Empty;
    }
}
