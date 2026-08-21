using System;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class NumberFormatInterpreter
    {
        public FormatRule Interpret(string numberFormat)
        {
            var format = numberFormat ?? string.Empty;
            var rule = new FormatRule
            {
                RawNumberFormat = format,
                IsPercentage = format.Contains("%"),
                IsScientificNotation = format.IndexOf("E+", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                       format.IndexOf("E-", StringComparison.OrdinalIgnoreCase) >= 0,
                UnitSuffix = ExtractUnitSuffix(format)
            };

            var decimalIndex = format.IndexOf('.');
            if (decimalIndex >= 0)
            {
                var decimals = 0;
                for (var i = decimalIndex + 1; i < format.Length; i++)
                {
                    if (format[i] == '0' || format[i] == '#')
                    {
                        decimals++;
                    }
                    else
                    {
                        break;
                    }
                }

                rule.DecimalPlaces = decimals;
            }
            else if (format.IndexOf('0') >= 0 || format.IndexOf('#') >= 0)
            {
                rule.DecimalPlaces = 0;
            }

            return rule;
        }

        private static string ExtractUnitSuffix(string numberFormat)
        {
            var format = numberFormat ?? string.Empty;
            if (string.IsNullOrWhiteSpace(format))
            {
                return string.Empty;
            }

            var section = format.Split(';')[0];
            var literals = new System.Collections.Generic.List<string>();
            var inQuote = false;
            var buffer = string.Empty;

            for (var index = 0; index < section.Length; index++)
            {
                var ch = section[index];
                if (ch == '"')
                {
                    if (inQuote && buffer.Length > 0)
                    {
                        literals.Add(buffer);
                    }

                    buffer = string.Empty;
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote)
                {
                    buffer += ch;
                }
            }

            if (section.Contains("%") && !literals.Contains("%"))
            {
                literals.Add("%");
            }

            var suffix = string.Join(string.Empty, literals)
                .Replace("+/-", string.Empty)
                .Replace("±", string.Empty)
                .Replace("<=", string.Empty)
                .Replace(">=", string.Empty)
                .Replace("<", string.Empty)
                .Replace(">", string.Empty)
                .Replace("≤", string.Empty)
                .Replace("≥", string.Empty)
                .Trim();

            return suffix;
        }
    }
}
