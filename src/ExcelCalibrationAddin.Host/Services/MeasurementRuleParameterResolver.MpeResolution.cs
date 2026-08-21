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
        private ResolvedMpe ResolveMpe(SheetSnapshot sheet, MeasurementRule rule)
        {
            if (sheet == null || rule?.MpeSource?.Range == null)
            {
                return null;
            }

            var logicalCells = MergedCellLogicalRangeResolver.GetTextCells(sheet, rule.MpeSource.Range);
            if (logicalCells.Count == 0)
            {
                return null;
            }

            var rangeContextTexts = CollectRangeContextTexts(sheet, rule.MpeSource.Range);

            var candidates = new List<MpeCandidate>();

            foreach (var logicalCell in logicalCells)
            {
                var contextTexts = new List<string>(rangeContextTexts);
                contextTexts.AddRange(CollectCellContextTexts(sheet, logicalCell));
                var candidate = BuildCandidate(logicalCell, contextTexts, sheet, rule);
                if (candidate != null)
                {
                    candidates.Add(candidate);
                }
            }

            var best = candidates
                .OrderByDescending(item => item.Score)
                .ThenBy(item => item.Row)
                .ThenBy(item => item.Column)
                .FirstOrDefault();

            if (best == null)
            {
                return null;
            }

            return new ResolvedMpe
            {
                ErrorType = best.ErrorType,
                Mpe = best.Mpe,
                NegativeTolerance = best.NegativeTolerance,
                PositiveTolerance = best.PositiveTolerance,
                ReferenceRange = best.ReferenceRange,
                RequirementOperator = best.RequirementOperator,
                ValuePattern = best.ValuePattern
            };
        }

        private MpeCandidate BuildCandidate(
            LogicalCellRange logicalCell,
            IReadOnlyCollection<string> contextTexts,
            SheetSnapshot sheet,
            MeasurementRule rule)
        {
            var anchor = logicalCell?.Anchor;
            if (anchor == null)
            {
                return null;
            }

            var requirement = RequirementTextParser.Parse(anchor);
            var rawText = requirement.Text;
            var numberFormat = anchor.NumberFormat ?? string.Empty;
            var unitSignals = string.Join(" ", contextTexts
                .Concat(requirement.ContextSignals)
                .Where(text => !string.IsNullOrWhiteSpace(text)));
            var templatePattern = MpeValuePatternCodec.Parse(rule?.MpeSource?.ValuePattern);
            var displayedErrorType = ResolveDisplayedErrorType(requirement, rawText, numberFormat, anchor.Formula);
            if (templatePattern != null &&
                displayedErrorType.HasValue &&
                templatePattern.ErrorType != displayedErrorType.Value)
            {
                templatePattern = null;
            }

            var errorType = displayedErrorType ?? templatePattern?.ErrorType ?? ResolveErrorType(unitSignals);
            var numberCandidates = ExtractNumberCandidates(rawText);
            if (numberCandidates.Count == 0)
            {
                return null;
            }

            var selectedNumber = SelectBestNumber(numberCandidates, errorType, rawText, unitSignals);
            if (!selectedNumber.HasValue)
            {
                return null;
            }

            var mpe = NormalizeMpeValue(selectedNumber.Value, errorType, rawText, numberFormat, unitSignals, templatePattern);
            if (double.IsNaN(mpe) || double.IsInfinity(mpe) || mpe <= 0)
            {
                return null;
            }

            var toleranceRange = ResolveToleranceRange(rawText, errorType, numberFormat, unitSignals, templatePattern);
            var valuePattern = templatePattern?.RawPattern ??
            MpeValuePatternCodec.Build(
                    errorType,
                    ResolveMpeScaleFactor(selectedNumber.Value, errorType, rawText, numberFormat, unitSignals),
                    requirement.Operator);

            double? referenceRange = null;
            if (errorType == ErrorType.Referenced)
            {
                referenceRange = ResolveReferenceRange(sheet, rule.RangeSource?.Range);
            }

            return new MpeCandidate
            {
                ErrorType = errorType,
                Mpe = mpe,
                NegativeTolerance = toleranceRange.negative,
                PositiveTolerance = toleranceRange.positive,
                ReferenceRange = referenceRange,
                RequirementOperator = requirement.Operator,
                ValuePattern = valuePattern,
                Row = logicalCell.Range.StartRow,
                Column = logicalCell.Range.StartColumn,
                Score = ScoreCandidate(rawText, numberFormat, unitSignals, errorType, referenceRange, requirement.Operator)
            };
        }

        private static ErrorType? ResolveDisplayedErrorType(
            ParsedTechnicalRequirement requirement,
            string rawText,
            string numberFormat,
            string formula)
        {
            var displayedSignals = string.Join(" ", (requirement?.ContextSignals ?? new List<string>())
                .Where(text => !string.IsNullOrWhiteSpace(text)));
            if (IsReferencedSignal(rawText) || IsReferencedSignal(displayedSignals))
            {
                return ErrorType.Referenced;
            }

            if (ContainsPercentToken(rawText) ||
                ContainsPercentToken(numberFormat) ||
                ContainsPercentToken(displayedSignals))
            {
                return ErrorType.Relative;
            }

            return (ContainsExplicitAbsoluteUnit(rawText) || IsConditionalRequirementFormula(formula)) &&
                NumberRegex.IsMatch(rawText ?? string.Empty)
                ? ErrorType.Absolute
                : (ErrorType?)null;
        }

        private static bool ContainsExplicitAbsoluteUnit(string text)
        {
            var normalized = NormalizeSignal(text)
                .Replace("Μ", "U")
                .Replace("µ", "U");
            return normalized.Contains("MOL/MOL") ||
                normalized.Contains("PPM") ||
                normalized.Contains("PPB") ||
                normalized.Contains("MG/") ||
                normalized.Contains("UG/") ||
                normalized.Contains("PA") ||
                normalized.Contains("MPA") ||
                normalized.Contains("KPA") ||
                normalized.Contains("°C");
        }

        private static bool IsConditionalRequirementFormula(string formula)
        {
            return Regex.IsMatch(
                formula ?? string.Empty,
                @"(?:^|[^A-Z])IF\s*\(",
                RegexOptions.IgnoreCase);
        }

        private static (double? negative, double? positive) ResolveToleranceRange(
            string rawText,
            ErrorType errorType,
            string numberFormat,
            string unitSignals,
            MpeValuePattern templatePattern)
        {
            var signedCandidates = ExtractSignedNumberCandidates(rawText);
            if (signedCandidates.Count < 2)
            {
                return (null, null);
            }

            var negative = signedCandidates
                .Where(candidate => candidate.Value < 0)
                .Select(candidate => (double?)NormalizeMpeValue(Math.Abs(candidate.Value), errorType, rawText, numberFormat, unitSignals, templatePattern))
                .FirstOrDefault();
            var positive = signedCandidates
                .Where(candidate => candidate.Value > 0 || candidate.HasExplicitPositiveSign)
                .Select(candidate => (double?)NormalizeMpeValue(Math.Abs(candidate.Value), errorType, rawText, numberFormat, unitSignals, templatePattern))
                .FirstOrDefault();

            return (negative, positive);
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

            for (var index = 0; index < section.Length; index++)
            {
                var ch = section[index];
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

                if (ch == '%' ||
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
                    ch == '\uFF08' ||
                    ch == '\uFF09')
                {
                    signals.Add(ch.ToString());
                }
            }

            return string.Join(string.Empty, signals).Trim();
        }
        private static TechnicalRequirementOperator ResolveRequirementOperator(string text)
        {
            var normalized = NormalizeSignal(text)
                .Replace("\uFF1C", "<")
                .Replace("\uFF1E", ">")
                .Replace("\u2264", "<=")
                .Replace("\u2266", "<=")
                .Replace("\u2265", ">=")
                .Replace("\u2267", ">=");

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
    }
}
