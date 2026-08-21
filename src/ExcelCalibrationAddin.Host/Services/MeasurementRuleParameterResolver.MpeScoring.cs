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
        private static List<SignedNumericCandidate> ExtractSignedNumberCandidates(string rawText)
        {
            var candidates = new List<SignedNumericCandidate>();
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return candidates;
            }

            foreach (Match match in NumberRegex.Matches(rawText))
            {
                if (!double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
                {
                    continue;
                }

                candidates.Add(new SignedNumericCandidate
                {
                    Value = value,
                    HasExplicitPositiveSign = match.Value.TrimStart().StartsWith("+", StringComparison.Ordinal)
                });
            }

            return candidates;
        }

        private static double ScoreCandidate(
            string rawText,
            string numberFormat,
            string unitSignals,
            ErrorType errorType,
            double? referenceRange,
            TechnicalRequirementOperator requirementOperator)
        {
            var score = 0;
            if (!string.IsNullOrWhiteSpace(rawText))
            {
                score += 80;
            }

            if (!string.IsNullOrWhiteSpace(numberFormat))
            {
                score += 10;
            }

            if (ContainsPercentToken(rawText))
            {
                score += 30;
            }

            if (ContainsPercentToken(numberFormat))
            {
                score += 15;
            }

            if (errorType == ErrorType.Referenced)
            {
                score += 40;
                if (referenceRange.HasValue && referenceRange.Value > 0)
                {
                    score += 30;
                }
            }
            else if (errorType == ErrorType.Relative)
            {
                score += 25;
            }

            if (!string.IsNullOrWhiteSpace(unitSignals))
            {
                score += 10;
            }

            if (requirementOperator != TechnicalRequirementOperator.None)
            {
                score += 20;
            }

            return score;
        }

        private static ErrorType ResolveErrorType(string context)
        {
            var normalized = NormalizeSignal(context);
            if (ReferencedMarkers.Any(marker => normalized.Contains(NormalizeSignal(marker))))
            {
                return ErrorType.Referenced;
            }

            return normalized.Contains("%")
                ? ErrorType.Relative
                : ErrorType.Absolute;
        }

        private static double NormalizeMpeValue(
            double rawValue,
            ErrorType errorType,
            string rawText,
            string numberFormat,
            string unitSignals,
            MpeValuePattern templatePattern = null)
        {
            if (templatePattern != null)
            {
                return Math.Abs(rawValue) * templatePattern.ScaleFactor;
            }

            var magnitude = Math.Abs(rawValue);
            if (errorType == ErrorType.Absolute)
            {
                return magnitude;
            }

            if (ContainsPercentToken(numberFormat) && magnitude <= 1)
            {
                return magnitude;
            }

            if (ContainsPercentToken(rawText))
            {
                return magnitude / 100.0;
            }

            if (ContainsPercentToken(numberFormat))
            {
                return magnitude;
            }

            if (ContainsPercentToken(unitSignals))
            {
                return magnitude > 1 ? magnitude / 100.0 : magnitude;
            }

            return magnitude;
        }

        private static double ResolveMpeScaleFactor(
            double rawValue,
            ErrorType errorType,
            string rawText,
            string numberFormat,
            string unitSignals)
        {
            if (errorType == ErrorType.Absolute)
            {
                return 1d;
            }

            var magnitude = Math.Abs(rawValue);
            if (ContainsPercentToken(numberFormat) && magnitude <= 1)
            {
                return 1d;
            }

            if (ContainsPercentToken(rawText))
            {
                return 0.01d;
            }

            if (ContainsPercentToken(numberFormat))
            {
                return 1d;
            }

            if (ContainsPercentToken(unitSignals))
            {
                return magnitude > 1 ? 0.01d : 1d;
            }

            return 1d;
        }

        private static double? SelectBestNumber(
            IReadOnlyList<NumericCandidate> numberCandidates,
            ErrorType errorType,
            string rawText,
            string unitSignals)
        {
            if (numberCandidates == null || numberCandidates.Count == 0)
            {
                return null;
            }

            IEnumerable<NumericCandidate> preferred = numberCandidates;
            switch (errorType)
            {
                case ErrorType.Referenced:
                    preferred = numberCandidates
                        .Where(item => ContainsPercentToken(item.Context) || IsReferencedSignal(item.Context))
                        .DefaultIfEmpty(numberCandidates[0]);
                    break;
                case ErrorType.Relative:
                    preferred = numberCandidates
                        .Where(item => ContainsPercentToken(item.Context) || ContainsPercentToken(rawText) || ContainsPercentToken(unitSignals))
                        .DefaultIfEmpty(numberCandidates[0]);
                    break;
                default:
                    preferred = numberCandidates
                        .Where(item => !ContainsPercentToken(item.Context))
                        .DefaultIfEmpty(numberCandidates[0]);
                    break;
            }

            return preferred
                .OrderByDescending(item => ScoreNumberCandidate(item, errorType))
                .ThenBy(item => item.Index)
                .Select(item => (double?)item.Value)
                .FirstOrDefault();
        }

        private static int ScoreNumberCandidate(NumericCandidate candidate, ErrorType errorType)
        {
            var score = 0;
            if (errorType == ErrorType.Absolute && !ContainsPercentToken(candidate.Context))
            {
                score += 20;
            }

            if (errorType != ErrorType.Absolute && ContainsPercentToken(candidate.Context))
            {
                score += 20;
            }

            if (errorType == ErrorType.Referenced && IsReferencedSignal(candidate.Context))
            {
                score += 20;
            }

            if (candidate.Value > 0)
            {
                score += 10;
            }

            return score;
        }

        private static bool IsReferencedSignal(string text)
        {
            var normalized = NormalizeSignal(text);
            return ReferencedMarkers.Any(marker => normalized.Contains(NormalizeSignal(marker)));
        }

        private static List<NumericCandidate> ExtractNumberCandidates(string rawText)
        {
            var candidates = new List<NumericCandidate>();
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return candidates;
            }

            var matches = NumberRegex.Matches(rawText);
            foreach (Match match in matches)
            {
                if (!double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
                {
                    continue;
                }

                var snippetStart = Math.Max(0, match.Index - 6);
                var snippetLength = Math.Min(rawText.Length - snippetStart, match.Length + 12);
                candidates.Add(new NumericCandidate
                {
                    Value = value,
                    Index = match.Index,
                    Context = rawText.Substring(snippetStart, snippetLength)
                });
            }

            return candidates;
        }

    }
}
