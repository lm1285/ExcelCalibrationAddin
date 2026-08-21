using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed partial class MeasurementValueGenerator
    {
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
                return NextTriangular(lower, upper, (lower + upper) / 2.0);
            }

            var mean = (lower + upper) / 2.0;
            var sigma = Math.Max((upper - lower) / 4.0, 1e-9);

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

        private double NextTriangular(double lower, double upper, double mode)
        {
            var sample = _random.NextDouble();
            var split = (mode - lower) / (upper - lower);
            if (sample < split)
            {
                return lower + Math.Sqrt(sample * (upper - lower) * (mode - lower));
            }

            return upper - Math.Sqrt((1 - sample) * (upper - lower) * (upper - mode));
        }

        private double NextCenteredOffset(double halfWidth)
        {
            return (_random.NextDouble() * 2.0 - 1.0) * halfWidth;
        }

        private double ApplyInstrumentResolution(double value, int decimalPlaces)
        {
            if (decimalPlaces < 0)
            {
                return value;
            }

            var resolution = Math.Pow(10, -decimalPlaces);
            if (resolution <= 0)
            {
                return value;
            }

            return Math.Round(value / resolution) * resolution;
        }

        private double ClampMagnitudeToConfiguredRange(
            double value,
            double minimumMagnitude,
            double maximumMagnitude,
            int decimalPlaces,
            double allowedMagnitude)
        {
            if (maximumMagnitude < minimumMagnitude)
            {
                maximumMagnitude = minimumMagnitude;
            }

            var clamped = Math.Max(minimumMagnitude, Math.Min(maximumMagnitude, value));
            if (decimalPlaces < 0)
            {
                return clamped;
            }

            var resolution = Math.Pow(10, -decimalPlaces);
            if (resolution <= 0)
            {
                return clamped;
            }

            var minimumStep = Math.Ceiling(minimumMagnitude / resolution) * resolution;
            var maximumStep = Math.Floor(maximumMagnitude / resolution) * resolution;
            if (minimumStep <= maximumStep)
            {
                return Math.Max(minimumStep, Math.Min(maximumStep, clamped));
            }

            var nearestNonZeroStep = Math.Ceiling(minimumMagnitude / resolution) * resolution;
            if (nearestNonZeroStep <= allowedMagnitude + 1e-12)
            {
                return nearestNonZeroStep;
            }

            return 0;
        }

        private (double minimum, double maximum) ResolveRangeCoefficients(MeasurementGenerationInput input, int direction)
        {
            if (IsMinimumRequirement(input.RequirementOperator))
            {
                return (
                    _configuration.MinimumRequirementMinimumCoefficient,
                    _configuration.MinimumRequirementMaximumCoefficient);
            }

            var coefficientOverride = input?.CoefficientOverride;
            if (!_configuration.UseIndependentDeviationControl)
            {
                return (
                    coefficientOverride?.PositiveMinimumCoefficient ??
                    coefficientOverride?.NegativeMinimumCoefficient ??
                    _configuration.UnifiedErrorMinimumCoefficient,
                    coefficientOverride?.PositiveMaximumCoefficient ??
                    coefficientOverride?.NegativeMaximumCoefficient ??
                    _configuration.UnifiedErrorMaximumCoefficient);
            }

            if (direction > 0)
            {
                return (
                    coefficientOverride?.PositiveMinimumCoefficient ?? _configuration.PositiveErrorMinimumCoefficient,
                    coefficientOverride?.PositiveMaximumCoefficient ?? _configuration.PositiveErrorMaximumCoefficient);
            }

            return (
                coefficientOverride?.NegativeMinimumCoefficient ?? _configuration.NegativeErrorMinimumCoefficient,
                coefficientOverride?.NegativeMaximumCoefficient ?? _configuration.NegativeErrorMaximumCoefficient);
        }

        private static int ResolveDecimalPlaces(MeasurementGenerationInput input, int index)
        {
            if (input?.DecimalPlacesByValue != null &&
                index >= 0 &&
                index < input.DecimalPlacesByValue.Count)
            {
                return input.DecimalPlacesByValue[index];
            }

            return input?.DecimalPlaces ?? 2;
        }

        private List<double> EnsureVisibleVariation(
            List<double> errors,
            MeasurementGenerationInput input,
            double lower,
            double upper,
            double maxSpread)
        {
            var result = new List<double>(errors.Count);
            for (var index = 0; index < errors.Count; index++)
            {
                var requested = errors[index];
                var direction = requested < 0 ? -1 : 1;
                var allowedMagnitude = ResolveAllowedMagnitude(input.StandardValue, lower, upper, direction);
                var coefficients = ResolveRangeCoefficients(input, direction);
                var magnitudeBounds = ResolveMagnitudeBounds(input, allowedMagnitude, coefficients);
                var minimumMagnitude = magnitudeBounds.minimum;
                var maximumMagnitude = magnitudeBounds.maximum;

                var decimalPlaces = ResolveDecimalPlaces(input, index);
                var candidate = direction * ClampMagnitudeToConfiguredRange(
                    ApplyInstrumentResolution(Math.Abs(requested), decimalPlaces),
                    minimumMagnitude,
                    maximumMagnitude,
                    decimalPlaces,
                    allowedMagnitude);
                if (Math.Abs(candidate) <= 1e-12)
                {
                    throw new InvalidOperationException(
                        "Unable to represent a non-zero calibration error with the configured resolution.");
                }

                result.Add(candidate);
            }

            if (input.ValueCount <= 1 ||
                HasVisibleVariation(result, input))
            {
                return result;
            }

            var startIndex = _random.Next(0, result.Count);
            var firstOffsetDirection = _random.Next(0, 2) == 0 ? -1 : 1;
            for (var candidateIndex = 0; candidateIndex < result.Count; candidateIndex++)
            {
                var index = (startIndex + candidateIndex) % result.Count;
                var requested = result[index];
                var direction = requested < 0 ? -1 : 1;
                var allowedMagnitude = ResolveAllowedMagnitude(input.StandardValue, lower, upper, direction);
                var coefficients = ResolveRangeCoefficients(input, direction);
                var magnitudeBounds = ResolveMagnitudeBounds(input, allowedMagnitude, coefficients);
                var decimalPlaces = ResolveDecimalPlaces(input, index);
                var resolution = Math.Pow(10, -decimalPlaces);

                foreach (var offsetDirection in new[] { firstOffsetDirection, -firstOffsetDirection })
                {
                    var magnitude = Math.Abs(requested) + (offsetDirection * resolution);
                    if (magnitude < magnitudeBounds.minimum - 1e-12 ||
                        magnitude > magnitudeBounds.maximum + 1e-12 ||
                        magnitude > allowedMagnitude + 1e-12)
                    {
                        continue;
                    }

                    var alternative = direction * ApplyInstrumentResolution(magnitude, decimalPlaces);
                    var adjusted = result.ToList();
                    adjusted[index] = alternative;
                    if (Math.Abs(alternative) > 1e-12 &&
                        IsSpreadValid(adjusted, maxSpread) &&
                        HasVisibleVariation(adjusted, input))
                    {
                        return adjusted;
                    }
                }
            }

            throw new InvalidOperationException(
                "Unable to generate visibly different repeated measurements within the configured error interval and measurement resolution.");
        }

        private static bool HasVisibleVariation(IReadOnlyList<double> errors, MeasurementGenerationInput input)
        {
            return errors
                .Select((error, index) => Math.Round(
                    input.StandardValue + error,
                    ResolveDecimalPlaces(input, index)))
                .Select(value => Math.Round(value, 15))
                .Distinct()
                .Count() > 1;
        }

        private static bool IsSpreadValid(List<double> values, double maxSpread)
        {
            var tolerance = Math.Max(1e-15, Math.Abs(maxSpread) * 1e-12);
            return values.Max() - values.Min() <= maxSpread + tolerance;
        }

    }
}
