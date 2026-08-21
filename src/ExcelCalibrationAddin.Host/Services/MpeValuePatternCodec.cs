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
            TechnicalRequirementOperator requirementOperator = TechnicalRequirementOperator.None)
        {
            var pattern = Prefix +
                errorType.ToString().ToLowerInvariant() +
                ":scale=" +
                scaleFactor.ToString("G17", CultureInfo.InvariantCulture);

            if (requirementOperator != TechnicalRequirementOperator.None)
            {
                pattern += ":op=" + requirementOperator.ToString().ToLowerInvariant();
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
            foreach (var part in parts.Skip(2))
            {
                if (!part.StartsWith("op=", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Enum.TryParse(part.Substring("op=".Length), true, out requirementOperator);
            }

            return new MpeValuePattern
            {
                ErrorType = errorType,
                ScaleFactor = scaleFactor,
                Operator = requirementOperator,
                RawPattern = valuePattern.Trim()
            };
        }
    }

    internal sealed class MpeValuePattern
    {
        public ErrorType ErrorType { get; set; }
        public double ScaleFactor { get; set; } = 1d;
        public TechnicalRequirementOperator Operator { get; set; }
        public string RawPattern { get; set; } = string.Empty;
    }
}
