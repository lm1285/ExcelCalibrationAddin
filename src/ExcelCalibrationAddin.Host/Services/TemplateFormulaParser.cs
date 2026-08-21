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
    internal static class TemplateFormulaParser
    {
        private static readonly Regex ReferenceRegex = new Regex(
            @"(?:(?:'(?<quotedSheet>[^']+)'|(?<sheet>[A-Za-z0-9_一-鿿]+))!)?(?<start>\$?[A-Z]{1,3}\$?\d+)(?::(?<end>\$?[A-Z]{1,3}\$?\d+))?",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex NumberRegex = new Regex(
            @"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?",
            RegexOptions.Compiled);

        public static TemplateFormulaDefinition Parse(
            SheetSnapshot sheet,
            CellMeta formulaCell,
            IReadOnlyDictionary<TemplateRegionRole, CellRange> roleRanges)
        {
            if (formulaCell == null || string.IsNullOrWhiteSpace(formulaCell.Formula))
            {
                return null;
            }

            var definition = new TemplateFormulaDefinition
            {
                Formula = formulaCell.Formula,
                FormulaR1C1 = formulaCell.FormulaR1C1,
                IsConditional = Regex.IsMatch(formulaCell.Formula, @"(?:^|[^A-Z])IF\s*\(", RegexOptions.IgnoreCase)
            };
            definition.References = ExtractReferences(
                formulaCell.Formula,
                sheet?.Name,
                roleRanges);

            if (definition.IsConditional)
            {
                ParseIfExpression(
                    sheet,
                    formulaCell.Formula,
                    string.Empty,
                    definition.Branches,
                    0);
            }

            definition.IsFullyParsed = !definition.IsConditional ||
                (definition.Branches.Count >= 2 && definition.Branches.All(HasResolvedBranchValue));
            return definition;
        }

        private static List<TemplateFormulaReference> ExtractReferences(
            string formula,
            string fallbackSheetName,
            IReadOnlyDictionary<TemplateRegionRole, CellRange> roleRanges)
        {
            var masked = MaskStringLiterals(formula);
            return ReferenceRegex.Matches(masked)
                .Cast<Match>()
                .Select(match => BuildReference(match, fallbackSheetName, roleRanges))
                .Where(reference => reference != null)
                .GroupBy(reference => reference.Token, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();
        }

        private static TemplateFormulaReference BuildReference(
            Match match,
            string fallbackSheetName,
            IReadOnlyDictionary<TemplateRegionRole, CellRange> roleRanges)
        {
            var start = ParseCellReference(match.Groups["start"].Value);
            var end = ParseCellReference(match.Groups["end"].Success
                ? match.Groups["end"].Value
                : match.Groups["start"].Value);
            if (start == null || end == null)
            {
                return null;
            }

            var sheetName = match.Groups["quotedSheet"].Success
                ? match.Groups["quotedSheet"].Value
                : match.Groups["sheet"].Success
                    ? match.Groups["sheet"].Value
                    : fallbackSheetName ?? string.Empty;
            var range = new CellRange
            {
                SheetName = sheetName,
                StartRow = Math.Min(start.Row, end.Row),
                EndRow = Math.Max(start.Row, end.Row),
                StartColumn = Math.Min(start.Column, end.Column),
                EndColumn = Math.Max(start.Column, end.Column)
            };

            var role = TemplateRegionRole.Unknown;
            foreach (var candidate in roleRanges ?? new Dictionary<TemplateRegionRole, CellRange>())
            {
                if (RangesOverlap(range, candidate.Value))
                {
                    role = candidate.Key;
                    break;
                }
            }

            return new TemplateFormulaReference
            {
                Token = match.Value,
                SheetName = sheetName,
                Range = range,
                Role = role
            };
        }

        private static void ParseIfExpression(
            SheetSnapshot sheet,
            string expression,
            string parentCondition,
            ICollection<TemplateFormulaBranch> branches,
            int depth)
        {
            if (depth > 8 || !TryParseIfArguments(expression, out var arguments))
            {
                AddBranch(sheet, parentCondition, expression, branches);
                return;
            }

            var condition = arguments[0].Trim();
            var trueCondition = CombineCondition(parentCondition, condition);
            if (LooksLikeIf(arguments[1]))
            {
                ParseIfExpression(sheet, arguments[1], trueCondition, branches, depth + 1);
            }
            else
            {
                AddBranch(sheet, trueCondition, arguments[1], branches);
            }

            var falseCondition = CombineCondition(parentCondition, "NOT(" + condition + ")");
            if (LooksLikeIf(arguments[2]))
            {
                ParseIfExpression(sheet, arguments[2], falseCondition, branches, depth + 1);
            }
            else
            {
                AddBranch(sheet, falseCondition, arguments[2], branches);
            }
        }

        private static void AddBranch(
            SheetSnapshot sheet,
            string condition,
            string expression,
            ICollection<TemplateFormulaBranch> branches)
        {
            var displayText = ResolveExpressionDisplayText(sheet, expression);
            var parsed = RequirementTextParser.Parse(new CellMeta
            {
                Text = displayText,
                DisplayText = displayText,
                RawValueText = displayText
            });
            var numericValue = 0d;
            var numberMatch = NumberRegex.Match(displayText ?? string.Empty);
            var hasNumber = numberMatch.Success && double.TryParse(
                numberMatch.Value,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out numericValue);

            branches.Add(new TemplateFormulaBranch
            {
                Condition = condition ?? string.Empty,
                ValueExpression = (expression ?? string.Empty).Trim(),
                DisplayText = displayText,
                NumericValue = hasNumber ? (double?)numericValue : null,
                Operator = parsed.Operator,
                Unit = TemplateUnitParser.Extract(displayText)
            });
        }

        private static string ResolveExpressionDisplayText(SheetSnapshot sheet, string expression)
        {
            var value = (expression ?? string.Empty).Trim();
            if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"')
            {
                return value.Substring(1, value.Length - 2).Replace("\"\"", "\"");
            }

            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
            {
                return value;
            }

            var match = ReferenceRegex.Match(value);
            if (!match.Success || match.Index != 0 || match.Length != value.Length)
            {
                return string.Empty;
            }

            var reference = BuildReference(
                match,
                sheet?.Name,
                new Dictionary<TemplateRegionRole, CellRange>());
            if (reference?.Range == null || sheet == null ||
                !string.Equals(reference.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            return sheet.Cells.FirstOrDefault(cell =>
                cell.Row == reference.Range.StartRow &&
                cell.Column == reference.Range.StartColumn)?.Text ?? string.Empty;
        }

        private static bool TryParseIfArguments(string expression, out List<string> arguments)
        {
            arguments = new List<string>();
            var value = (expression ?? string.Empty).Trim().TrimStart('=');
            if (!LooksLikeIf(value))
            {
                return false;
            }

            var open = value.IndexOf('(');
            var close = FindMatchingParenthesis(value, open);
            if (open < 0 || close < 0)
            {
                return false;
            }

            arguments = SplitArguments(value.Substring(open + 1, close - open - 1));
            return arguments.Count == 3;
        }

        private static List<string> SplitArguments(string value)
        {
            var result = new List<string>();
            var start = 0;
            var depth = 0;
            var inString = false;
            for (var index = 0; index < value.Length; index++)
            {
                var ch = value[index];
                if (ch == '"')
                {
                    if (inString && index + 1 < value.Length && value[index + 1] == '"')
                    {
                        index++;
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                if (inString) continue;
                if (ch == '(') depth++;
                else if (ch == ')') depth--;
                else if ((ch == ',' || ch == ';') && depth == 0)
                {
                    result.Add(value.Substring(start, index - start));
                    start = index + 1;
                }
            }

            result.Add(value.Substring(start));
            return result;
        }

        private static int FindMatchingParenthesis(string value, int open)
        {
            if (open < 0) return -1;
            var depth = 0;
            var inString = false;
            for (var index = open; index < value.Length; index++)
            {
                if (value[index] == '"') inString = !inString;
                if (inString) continue;
                if (value[index] == '(') depth++;
                else if (value[index] == ')' && --depth == 0) return index;
            }

            return -1;
        }

        private static bool LooksLikeIf(string expression)
        {
            return Regex.IsMatch(
                (expression ?? string.Empty).Trim().TrimStart('='),
                @"^IF\s*\(",
                RegexOptions.IgnoreCase);
        }

        private static string CombineCondition(string parent, string current)
        {
            return string.IsNullOrWhiteSpace(parent)
                ? current
                : "(" + parent + ") AND (" + current + ")";
        }

        private static bool HasResolvedBranchValue(TemplateFormulaBranch branch)
        {
            return branch != null &&
                (!string.IsNullOrWhiteSpace(branch.DisplayText) || branch.NumericValue.HasValue);
        }

        private static string MaskStringLiterals(string value)
        {
            var builder = new StringBuilder(value ?? string.Empty);
            var inString = false;
            for (var index = 0; index < builder.Length; index++)
            {
                if (builder[index] == '"')
                {
                    inString = !inString;
                    continue;
                }

                if (inString) builder[index] = ' ';
            }

            return builder.ToString();
        }

        private static CellAddress ParseCellReference(string value)
        {
            var normalized = (value ?? string.Empty).Replace("$", string.Empty).ToUpperInvariant();
            var letters = new string(normalized.TakeWhile(char.IsLetter).ToArray());
            var digits = new string(normalized.SkipWhile(char.IsLetter).ToArray());
            if (letters.Length == 0 || !int.TryParse(digits, out var row)) return null;
            var column = 0;
            foreach (var ch in letters) column = column * 26 + ch - 'A' + 1;
            return row > 0 && column > 0
                ? new CellAddress { Row = row, Column = column }
                : null;
        }

        private static bool RangesOverlap(CellRange left, CellRange right)
        {
            return left != null && right != null &&
                string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase) &&
                left.StartRow <= right.EndRow && left.EndRow >= right.StartRow &&
                left.StartColumn <= right.EndColumn && left.EndColumn >= right.StartColumn;
        }
    }
}
