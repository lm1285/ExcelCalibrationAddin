using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public sealed partial class GenerateMeasurementUseCase
    {
        private bool CanAddTrendError(
            MeasurementRule rule,
            double standardValue,
            double trendError,
            IEnumerable<TrendErrorSample> existing)
        {
            var usage = ResolveTrendErrorUsage(rule, standardValue, trendError);
            if (!usage.HasValue)
            {
                return true;
            }

            var minimumChange = _generationConfiguration.ResultGroupMinimumFluctuationCoefficient;
            var maximumChange = _generationConfiguration.ResultGroupMaximumFluctuationCoefficient;
            var tolerance = Math.Max(1e-12, Math.Abs(maximumChange) * 1e-12);
            return existing.All(sample =>
                !sample.Usage.HasValue ||
                (Math.Abs(usage.Value - sample.Usage.Value) >= minimumChange - tolerance &&
                 Math.Abs(usage.Value - sample.Usage.Value) <= maximumChange + tolerance));
        }

        private bool TryAddTrendError(
            MeasurementRule rule,
            double standardValue,
            double trendError,
            ICollection<TrendErrorSample> existing)
        {
            if (Math.Abs(trendError) <= 1e-12)
            {
                return false;
            }

            var sameStandard = existing.FirstOrDefault(sample =>
                Math.Abs(sample.StandardValue - standardValue) <= 1e-12);
            if (sameStandard != null)
            {
                var usage = ResolveTrendErrorUsage(rule, standardValue, trendError);
                return usage.HasValue && sameStandard.Usage.HasValue
                    ? usage.Value.Equals(sameStandard.Usage.Value)
                    : sameStandard.Value.Equals(trendError);
            }

            if (HasDuplicateTrendError(standardValue, trendError, existing) ||
                !CanAddTrendError(rule, standardValue, trendError, existing))
            {
                return false;
            }

            AddTrendError(rule, standardValue, trendError, existing);
            return true;
        }

        private static bool HasDuplicateTrendError(
            double standardValue,
            double trendError,
            IEnumerable<TrendErrorSample> existing)
        {
            return existing.Any(sample =>
                Math.Abs(sample.StandardValue - standardValue) > 1e-12 &&
                sample.Value.Equals(trendError));
        }

        private static void AddTrendError(
            MeasurementRule rule,
            double standardValue,
            double trendError,
            ICollection<TrendErrorSample> existing)
        {
            existing.Add(new TrendErrorSample(
                standardValue,
                trendError,
                ResolveTrendErrorUsage(rule, standardValue, trendError)));
        }

        private static double? ResolveTrendErrorUsage(
            MeasurementRule rule,
            double standardValue,
            double trendError)
        {
            var allowed = ResolveAllowedFormulaErrorMagnitude(rule, standardValue);
            return allowed.HasValue && allowed.Value > 0
                ? trendError / allowed.Value
                : (double?)null;
        }

        private sealed class GeneratedPoint
        {
            public GeneratedPoint(MeasurementGenerationResult result, double representativeError)
            {
                Result = result;
                RepresentativeError = representativeError;
            }

            public MeasurementGenerationResult Result { get; }
            public double RepresentativeError { get; }
        }

        private sealed class TrendErrorSample
        {
            public TrendErrorSample(double standardValue, double value, double? usage)
            {
                StandardValue = standardValue;
                Value = value;
                Usage = usage;
            }

            public double StandardValue { get; }
            public double Value { get; }
            public double? Usage { get; }
        }
    }
}
