using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Services
{
    internal static partial class TemplateFormulaParser
    {
        private static readonly Regex ComparisonRegex = new Regex(
            @"^(?<left>.*?)(?<operator><=|>=|<>|=|<|>)(?<right>.*)$",
            RegexOptions.Compiled);

        public static bool TryResolveActiveValueExpression(
            SheetSnapshot sheet,
            string formula,
            out string expression)
        {
            expression = null;
            var value = (formula ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (!LooksLikeIf(value.TrimStart('=')) &&
                !Regex.IsMatch(value, @"(?:^|[^A-Z])IF\s*\(", RegexOptions.IgnoreCase))
            {
                return false;
            }

            return TryResolveActiveValueExpressionCore(sheet, value, 0, out expression);
        }

        public static IReadOnlyList<CellRange> ExtractReferencedRanges(string formula, string fallbackSheetName)
        {
            var ranges = new List<CellRange>();
            if (string.IsNullOrWhiteSpace(formula))
            {
                return ranges;
            }

            foreach (Match match in ReferenceRegex.Matches(MaskStringLiterals(formula)))
            {
                if (IsExternalWorkbookReference(formula, match))
                {
                    continue;
                }

                var reference = BuildReference(
                    match,
                    fallbackSheetName,
                    new Dictionary<TemplateRegionRole, CellRange>());
                if (reference?.Range == null)
                {
                    continue;
                }

                ranges.Add(reference.Range);
            }

            return ranges;
        }

        private static bool TryResolveActiveValueExpressionCore(
            SheetSnapshot sheet,
            string expression,
            int depth,
            out string valueExpression)
        {
            valueExpression = null;
            if (depth > 8 || !TryParseIfArguments(expression, out var arguments))
            {
                return false;
            }

            object conditionValue;
            if (!TryEvaluate(sheet, arguments[0], out conditionValue) || !(conditionValue is bool))
            {
                return false;
            }

            var selected = (bool)conditionValue ? arguments[1] : arguments[2];
            if (LooksLikeIf(selected))
            {
                return TryResolveActiveValueExpressionCore(sheet, selected, depth + 1, out valueExpression);
            }

            valueExpression = (selected ?? string.Empty).Trim();
            return !string.IsNullOrWhiteSpace(valueExpression);
        }

        private static bool TryEvaluate(SheetSnapshot sheet, string expression, out object value)
        {
            value = null;
            var text = (expression ?? string.Empty).Trim().TrimStart('=');
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            if (TryEvaluateNot(sheet, text, out value) ||
                TryEvaluateAnd(sheet, text, out value) ||
                TryEvaluateCountIf(sheet, text, out value))
            {
                return true;
            }

            var comparison = ComparisonRegex.Match(text);
            if (comparison.Success &&
                FindOperatorAtDepthZero(text) >= 0)
            {
                object left;
                object right;
                if (!TryEvaluate(sheet, comparison.Groups["left"].Value, out left) ||
                    !TryEvaluate(sheet, comparison.Groups["right"].Value, out right))
                {
                    return false;
                }

                value = Compare(left, comparison.Groups["operator"].Value, right);
                return value is bool;
            }

            if (text.Length >= 2 && text[0] == '"' && text[text.Length - 1] == '"')
            {
                value = text.Substring(1, text.Length - 2).Replace("\"\"", "\"");
                return true;
            }

            double number;
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
            {
                value = number;
                return true;
            }

            var referenceMatch = ReferenceRegex.Match(text);
            if (referenceMatch.Success && referenceMatch.Index == 0 && referenceMatch.Length == text.Length)
            {
                return TryReadCellValue(sheet, referenceMatch, out value);
            }

            return false;
        }

        private static bool TryEvaluateNot(SheetSnapshot sheet, string text, out object value)
        {
            value = null;
            if (!Regex.IsMatch(text, @"^NOT\s*\(", RegexOptions.IgnoreCase))
            {
                return false;
            }

            var open = text.IndexOf('(');
            var close = FindMatchingParenthesis(text, open);
            if (open < 0 || close < 0 || close != text.Length - 1)
            {
                return false;
            }

            object inner;
            if (!TryEvaluate(sheet, text.Substring(open + 1, close - open - 1), out inner) || !(inner is bool))
            {
                return false;
            }

            value = !(bool)inner;
            return true;
        }

        private static bool TryEvaluateAnd(SheetSnapshot sheet, string text, out object value)
        {
            value = null;
            if (!Regex.IsMatch(text, @"^AND\s*\(", RegexOptions.IgnoreCase))
            {
                return false;
            }

            var open = text.IndexOf('(');
            var close = FindMatchingParenthesis(text, open);
            if (open < 0 || close < 0 || close != text.Length - 1)
            {
                return false;
            }

            var arguments = SplitArguments(text.Substring(open + 1, close - open - 1));
            if (arguments.Count == 0)
            {
                return false;
            }

            var result = true;
            foreach (var argument in arguments)
            {
                object argumentValue;
                if (!TryEvaluate(sheet, argument, out argumentValue) || !(argumentValue is bool))
                {
                    return false;
                }

                result = result && (bool)argumentValue;
            }

            value = result;
            return true;
        }

        private static bool TryEvaluateCountIf(SheetSnapshot sheet, string text, out object value)
        {
            value = null;
            if (!Regex.IsMatch(text, @"^COUNTIF\s*\(", RegexOptions.IgnoreCase))
            {
                return false;
            }

            var open = text.IndexOf('(');
            var close = FindMatchingParenthesis(text, open);
            if (open < 0 || close < 0 || close != text.Length - 1)
            {
                return false;
            }

            var arguments = SplitArguments(text.Substring(open + 1, close - open - 1));
            if (arguments.Count != 2)
            {
                return false;
            }

            var rangeMatch = ReferenceRegex.Match(arguments[0].Trim());
            if (!rangeMatch.Success)
            {
                return false;
            }

            var range = BuildReference(
                rangeMatch,
                sheet?.Name,
                new Dictionary<TemplateRegionRole, CellRange>())?.Range;
            object criteria;
            if (range == null || !TryEvaluate(sheet, arguments[1], out criteria))
            {
                return false;
            }

            var count = 0;
            foreach (var cell in sheet?.Cells ?? new List<CellMeta>())
            {
                if (cell.Row < range.StartRow ||
                    cell.Row > range.EndRow ||
                    cell.Column < range.StartColumn ||
                    cell.Column > range.EndColumn)
                {
                    continue;
                }

                object cellValue;
                if (TryReadRawCellValue(cell, out cellValue) && ValuesEqual(cellValue, criteria))
                {
                    count++;
                }
            }

            value = (double)count;
            return true;
        }

        private static bool TryReadCellValue(SheetSnapshot sheet, Match match, out object value)
        {
            value = null;
            var reference = BuildReference(
                match,
                sheet?.Name,
                new Dictionary<TemplateRegionRole, CellRange>());
            if (reference?.Range == null || sheet == null)
            {
                return false;
            }

            var cell = sheet.Cells.FirstOrDefault(item =>
                item.Row == reference.Range.StartRow &&
                item.Column == reference.Range.StartColumn);
            return TryReadRawCellValue(cell, out value);
        }

        private static bool TryReadRawCellValue(CellMeta cell, out object value)
        {
            value = null;
            if (cell == null)
            {
                return false;
            }

            var candidates = new[] { cell.RawValueText, cell.DisplayText, cell.Text };
            foreach (var candidate in candidates)
            {
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    continue;
                }

                var trimmed = candidate.Trim();
                double number;
                if (double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number) ||
                    double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out number))
                {
                    value = number;
                    return true;
                }

                value = trimmed;
                return true;
            }

            return false;
        }

        private static object Compare(object left, string comparisonOperator, object right)
        {
            double leftNumber;
            double rightNumber;
            if (TryConvertNumber(left, out leftNumber) && TryConvertNumber(right, out rightNumber))
            {
                switch (comparisonOperator)
                {
                    case "<": return leftNumber < rightNumber;
                    case "<=": return leftNumber <= rightNumber;
                    case ">": return leftNumber > rightNumber;
                    case ">=": return leftNumber >= rightNumber;
                    case "<>": return Math.Abs(leftNumber - rightNumber) > 1e-12;
                    default: return Math.Abs(leftNumber - rightNumber) <= 1e-12;
                }
            }

            var leftText = Convert.ToString(left, CultureInfo.InvariantCulture) ?? string.Empty;
            var rightText = Convert.ToString(right, CultureInfo.InvariantCulture) ?? string.Empty;
            var equal = string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase);
            switch (comparisonOperator)
            {
                case "<>": return !equal;
                case "=": return equal;
                default: return false;
            }
        }

        private static bool TryConvertNumber(object value, out double number)
        {
            if (value is double)
            {
                number = (double)value;
                return true;
            }

            return double.TryParse(
                Convert.ToString(value, CultureInfo.InvariantCulture),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out number);
        }

        private static bool ValuesEqual(object left, object right)
        {
            double leftNumber;
            double rightNumber;
            if (TryConvertNumber(left, out leftNumber) && TryConvertNumber(right, out rightNumber))
            {
                return Math.Abs(leftNumber - rightNumber) <= 1e-12;
            }

            return string.Equals(
                Convert.ToString(left, CultureInfo.InvariantCulture),
                Convert.ToString(right, CultureInfo.InvariantCulture),
                StringComparison.OrdinalIgnoreCase);
        }

        private static int FindOperatorAtDepthZero(string text)
        {
            var depth = 0;
            var inString = false;
            for (var index = 0; index < text.Length; index++)
            {
                var ch = text[index];
                if (ch == '"')
                {
                    inString = !inString;
                    continue;
                }

                if (inString)
                {
                    continue;
                }

                if (ch == '(')
                {
                    depth++;
                    continue;
                }

                if (ch == ')')
                {
                    depth--;
                    continue;
                }

                if (depth != 0)
                {
                    continue;
                }

                if (index + 1 < text.Length &&
                    ((ch == '<' && (text[index + 1] == '=' || text[index + 1] == '>')) ||
                     (ch == '>' && text[index + 1] == '=')))
                {
                    return index;
                }

                if (ch == '<' || ch == '>' || ch == '=')
                {
                    return index;
                }
            }

            return -1;
        }
    }
}
