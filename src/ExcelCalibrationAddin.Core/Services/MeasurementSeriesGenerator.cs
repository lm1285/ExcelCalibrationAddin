using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class MeasurementSeriesGenerator
    {
        private readonly Random _random;

        public MeasurementSeriesGenerator(Random random)
        {
            _random = random ?? throw new ArgumentNullException(nameof(random));
        }

        public List<double> GenerateRepeatabilityValues(
            MeasurementRule rule,
            double standardValue,
            double centerValue,
            double toleranceRatio,
            int valueCount,
            int decimalPlaces,
            GenerationConfiguration configuration,
            double minimumVisibleSpread = 0)
        {
            var resolution = ResolveResolution(decimalPlaces);
            var maxSpread = Math.Abs(standardValue) * Math.Max(toleranceRatio, resolution / Math.Max(Math.Abs(standardValue), 1d));
            var coefficients = ResolveAbsoluteRangeCoefficients(rule, configuration);
            var minimumSpread = Math.Max(
                Math.Max(resolution, Math.Min(maxSpread, Math.Abs(minimumVisibleSpread))),
                maxSpread * coefficients.minimum);
            var maximumSpread = Math.Max(minimumSpread, maxSpread * coefficients.maximum);
            var targetSpread = Math.Min(maximumSpread, Math.Max(minimumSpread, maxSpread * ResolveMidCoefficient(configuration)));
            var start = centerValue - (targetSpread / 2.0);
            var values = new List<double>(valueCount);

            for (var index = 0; index < valueCount; index++)
            {
                var ratio = valueCount <= 1 ? 0.5 : index / (double)(valueCount - 1);
                var jitter = valueCount <= 2 ? 0 : NextCenteredOffset(targetSpread / 8.0);
                values.Add(RoundToResolution(start + (targetSpread * ratio) + jitter, decimalPlaces));
            }

            if (!HasVariation(values, decimalPlaces) && valueCount > 1)
            {
                var candidate = RoundToResolution(values[0] + resolution, decimalPlaces);
                var candidateSpread = Math.Max(values.Max(), candidate) - Math.Min(values.Min(), candidate);
                if (candidateSpread <= maximumSpread + 1e-12)
                {
                    values[valueCount - 1] = candidate;
                }
            }

            if (valueCount > 1 && !HasVariation(values, decimalPlaces))
            {
                throw new InvalidOperationException(
                    "Unable to generate non-zero repeatability at the configured measurement resolution.");
            }

            return values;
        }

        public List<double> GenerateUpperLimitValues(
            MeasurementRule rule,
            double upperLimit,
            int valueCount,
            int decimalPlaces,
            GenerationConfiguration configuration)
        {
            var coefficients = ResolveAbsoluteRangeCoefficients(rule, configuration);
            var minimum = Math.Max(0, upperLimit * coefficients.minimum);
            var maximum = Math.Min(upperLimit, Math.Max(minimum, upperLimit * coefficients.maximum));
            var values = new List<double>(valueCount);

            for (var index = 0; index < valueCount; index++)
            {
                values.Add(RoundToResolution(NextMagnitude(ResolveDistributionMode(configuration), minimum, maximum), decimalPlaces));
            }

            return values;
        }

        public List<double> GenerateResponseTimeValues(
            double standardValue,
            double mpe,
            int valueCount,
            int decimalPlaces,
            GenerationConfiguration configuration,
            double? measurementLowerBound = null,
            double? measurementUpperBound = null)
        {
            var upperLimit = Math.Abs(mpe);
            var normalizedConfiguration = new GenerationConfigurationStore().Normalize(configuration);
            var maximumSpread = upperLimit <= normalizedConfiguration.ResponseTimeThresholdSeconds
                ? normalizedConfiguration.ResponseTimeBelowThresholdMaximumDifferenceSeconds
                : normalizedConfiguration.ResponseTimeAboveThresholdMaximumDifferenceSeconds;
            var center = Math.Max(0d, standardValue);
            if (measurementLowerBound.HasValue)
            {
                center = Math.Max(center, measurementLowerBound.Value);
            }
            if (measurementUpperBound.HasValue)
            {
                center = Math.Min(center, measurementUpperBound.Value);
            }

            var minimum = Math.Max(0d, center - (maximumSpread / 2d));
            var maximum = center + (maximumSpread / 2d);
            if (measurementLowerBound.HasValue)
            {
                minimum = Math.Max(minimum, measurementLowerBound.Value);
            }
            if (measurementUpperBound.HasValue)
            {
                maximum = Math.Min(maximum, measurementUpperBound.Value);
            }
            var values = new List<double>(valueCount);

            for (var index = 0; index < valueCount; index++)
            {
                values.Add(RoundToResolution(
                    NextMagnitude(ResolveDistributionMode(configuration), minimum, maximum),
                    decimalPlaces));
            }

            return values;
        }

        public static double ResolveRepeatabilityTolerance(double rawTolerance)
        {
            var tolerance = ResolveRatioTolerance(rawTolerance);
            if (tolerance <= 0)
            {
                return 0.02;
            }

            return tolerance > 0.1 ? 0.02 : tolerance;
        }

        public static double ResolveUpperLimit(MeasurementRule rule)
        {
            if (rule?.MeasurementUpperBound.HasValue == true)
            {
                return rule.MeasurementUpperBound.Value;
            }

            return Math.Abs(rule?.FixedMpe ?? 0);
        }

        public static List<string> FormatValues(IEnumerable<double> values, int decimalPlaces)
        {
            return values
                .Select(value => Math.Round(value, decimalPlaces).ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture))
                .ToList();
        }

        private static (double minimum, double maximum) ResolveAbsoluteRangeCoefficients(
            MeasurementRule rule,
            GenerationConfiguration configuration)
        {
            var coefficientOverride = rule?.GenerationCoefficientOverride;
            return (
                coefficientOverride?.AbsoluteMinimumCoefficient ?? configuration.AbsoluteErrorMinimumCoefficient,
                coefficientOverride?.AbsoluteMaximumCoefficient ?? configuration.AbsoluteErrorMaximumCoefficient);
        }

        private static double ResolveMidCoefficient(GenerationConfiguration configuration)
        {
            return Math.Max(0.2, Math.Min(0.8, (configuration.AbsoluteErrorMinimumCoefficient + configuration.AbsoluteErrorMaximumCoefficient) / 2.0));
        }

        private double NextMagnitude(DistributionMode mode, double lower, double upper)
        {
            if (upper <= lower)
            {
                return lower;
            }

            if (mode == DistributionMode.Uniform)
            {
                return lower + (_random.NextDouble() * (upper - lower));
            }

            if (mode == DistributionMode.Triangular)
            {
                var sample = _random.NextDouble();
                var midpoint = (lower + upper) / 2.0;
                var split = (midpoint - lower) / (upper - lower);
                if (sample < split)
                {
                    return lower + Math.Sqrt(sample * (upper - lower) * (midpoint - lower));
                }

                return upper - Math.Sqrt((1 - sample) * (upper - lower) * (upper - midpoint));
            }

            var mean = (lower + upper) / 2.0;
            var sigma = Math.Max((upper - lower) / 6.0, 1e-9);
            for (var index = 0; index < 20; index++)
            {
                var sample = NextNormal(mean, sigma);
                if (sample >= lower && sample <= upper)
                {
                    return sample;
                }
            }

            return Math.Min(upper, Math.Max(lower, mean));
        }

        private double NextNormal(double mean, double sigma)
        {
            var u1 = 1.0 - _random.NextDouble();
            var u2 = 1.0 - _random.NextDouble();
            var randStdNormal = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2);
            return mean + sigma * randStdNormal;
        }

        private double NextCenteredOffset(double halfWidth)
        {
            return (_random.NextDouble() * 2.0 - 1.0) * halfWidth;
        }

        private static bool HasVariation(IEnumerable<double> values, int decimalPlaces)
        {
            var valueList = values?.ToList() ?? new List<double>();
            if (valueList.Count <= 1)
            {
                return true;
            }

            return valueList
                .Select(value => Math.Round(value, decimalPlaces))
                .Distinct()
                .Count() > 1;
        }

        private static double ResolveRatioTolerance(double rawTolerance)
        {
            var value = Math.Abs(rawTolerance);
            return value > 1 ? value / 100.0 : value;
        }

        private static double ResolveResolution(int decimalPlaces)
        {
            return decimalPlaces < 0 ? 0 : Math.Pow(10, -decimalPlaces);
        }

        private static double RoundToResolution(double value, int decimalPlaces)
        {
            return Math.Round(value, decimalPlaces);
        }

        private static DistributionMode ResolveDistributionMode(GenerationConfiguration configuration)
        {
            var value = configuration?.DefaultDistribution;
            if (Enum.TryParse(value, true, out DistributionMode mode))
            {
                return mode;
            }

            return DistributionMode.Normal;
        }
    }
}
