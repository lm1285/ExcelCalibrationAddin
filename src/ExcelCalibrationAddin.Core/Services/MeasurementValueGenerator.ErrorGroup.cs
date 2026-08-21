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
        private List<double> GenerateErrorGroup(MeasurementGenerationInput input, double lower, double upper, int direction)
        {
            var allowedMagnitude = ResolveAllowedMagnitude(input.StandardValue, lower, upper, direction);
            var minimumSpread = input.ValueCount <= 1
                ? 0
                : allowedMagnitude * _configuration.MeasurementGroupMinimumFluctuationCoefficient;
            var maxSpread = input.ValueCount <= 1
                ? 0
                : allowedMagnitude * _configuration.MeasurementGroupMaximumFluctuationCoefficient;
            var effectiveMaxSpread = ResolveEffectiveMaximumSpread(input, maxSpread);
            var rangeCoefficients = ResolveRangeCoefficients(input, direction);
            var magnitudeBounds = ResolveMagnitudeBounds(input, allowedMagnitude, rangeCoefficients);
            var minimumMagnitude = magnitudeBounds.minimum;
            var maximumMagnitude = magnitudeBounds.maximum;
            var resolutionBuffer = input.ValueCount <= 1
                ? 0
                : (input.DecimalPlacesByValue ?? new List<int>())
                    .DefaultIfEmpty(input.DecimalPlaces)
                    .Select(value => Math.Pow(10, -value))
                    .Max() * 2.0;
            var samplingMinimumSpread = Math.Min(effectiveMaxSpread, minimumSpread + resolutionBuffer);
            var samplingMaximumSpread = effectiveMaxSpread;
            samplingMinimumSpread = Math.Min(samplingMaximumSpread, samplingMinimumSpread);
            var targetSpread = input.ValueCount <= 1
                ? 0
                : NextMagnitude(input.DistributionMode, samplingMinimumSpread, samplingMaximumSpread);
            targetSpread = Math.Min(targetSpread, Math.Max(0, maximumMagnitude - minimumMagnitude));
            var halfSpread = targetSpread / 2.0;
            var centerMinimum = minimumMagnitude + halfSpread;
            var centerMaximum = maximumMagnitude - halfSpread;
            if (centerMaximum < centerMinimum)
            {
                centerMinimum = minimumMagnitude;
                centerMaximum = maximumMagnitude;
                targetSpread = Math.Max(0, maximumMagnitude - minimumMagnitude);
                halfSpread = targetSpread / 2.0;
            }

            var baseMagnitude = ResolveBaseMagnitude(input, centerMinimum, centerMaximum, allowedMagnitude);
            baseMagnitude = ClampMagnitudeToConfiguredRange(
                ApplyInstrumentResolution(baseMagnitude, ResolveDecimalPlaces(input, 0)),
                centerMinimum,
                centerMaximum,
                ResolveDecimalPlaces(input, 0),
                allowedMagnitude);

            var offsets = BuildCenteredOffsets(input.ValueCount, targetSpread);
            var errors = new List<double>(input.ValueCount);
            for (var index = 0; index < input.ValueCount; index++)
            {
                var valueDirection = _configuration.UseSameDeviationDirection
                    ? direction
                    : ResolveDirection(input, lower, upper);
                var valueAllowedMagnitude = ResolveAllowedMagnitude(input.StandardValue, lower, upper, valueDirection);
                var valueCoefficients = ResolveRangeCoefficients(input, valueDirection);
                var valueMagnitudeBounds = ResolveMagnitudeBounds(input, valueAllowedMagnitude, valueCoefficients);
                var valueMinimumMagnitude = valueMagnitudeBounds.minimum;
                var valueMaximumMagnitude = valueMagnitudeBounds.maximum;

                var decimalPlaces = ResolveDecimalPlaces(input, index);
                var candidateMagnitude = ClampMagnitudeToConfiguredRange(
                    ApplyInstrumentResolution(baseMagnitude + offsets[index], decimalPlaces),
                    valueMinimumMagnitude,
                    valueMaximumMagnitude,
                    decimalPlaces,
                    valueAllowedMagnitude);
                errors.Add(valueDirection * candidateMagnitude);
            }

            if (!IsSpreadValid(errors, effectiveMaxSpread))
            {
                var coarsestResolutionIndex = Enumerable.Range(0, errors.Count)
                    .OrderBy(index => ResolveDecimalPlaces(input, index))
                    .First();
                errors = Enumerable.Repeat(errors[coarsestResolutionIndex], errors.Count).ToList();
            }

            return EnsureVisibleVariation(errors, input, lower, upper, effectiveMaxSpread);
        }

        private double ResolveEffectiveMaximumSpread(MeasurementGenerationInput input, double configuredMaxSpread)
        {
            if (input.ValueCount <= 1)
            {
                return configuredMaxSpread;
            }

            var coarsestResolution = (input.DecimalPlacesByValue ?? new List<int>())
                .DefaultIfEmpty(input.DecimalPlaces)
                .Select(value => Math.Pow(10, -value))
                .Max();
            return Math.Max(configuredMaxSpread, coarsestResolution);
        }

        private List<double> BuildCenteredOffsets(int valueCount, double targetSpread)
        {
            if (valueCount <= 1 || targetSpread <= 0)
            {
                return Enumerable.Repeat(0d, Math.Max(1, valueCount)).Take(valueCount).ToList();
            }

            var offsets = new List<double>(valueCount);
            var spacing = targetSpread / (valueCount - 1);
            for (var index = 0; index < valueCount; index++)
            {
                var offset = (-targetSpread / 2.0) + (spacing * index);
                if (index > 0 && index < valueCount - 1)
                {
                    offset += NextCenteredOffset(spacing * 0.2);
                }

                offsets.Add(offset);
            }

            var average = offsets.Average();
            return offsets.Select(value => value - average).ToList();
        }

        private static double ResolveRequirementMagnitude(MeasurementGenerationInput input)
        {
            switch (input.ErrorType)
            {
                case ErrorType.Relative:
                    return Math.Abs(input.StandardValue) * input.Mpe;
                case ErrorType.Referenced:
                    return Math.Abs(input.ReferenceRange.GetValueOrDefault()) * input.Mpe;
                default:
                    return input.Mpe;
            }
        }

        private static (double minimum, double maximum) ResolveMagnitudeBounds(
            MeasurementGenerationInput input,
            double allowedMagnitude,
            (double minimum, double maximum) coefficients)
        {
            if (IsMinimumRequirement(input.RequirementOperator))
            {
                var requirementMagnitude = ResolveRequirementMagnitude(input);
                var minimum = requirementMagnitude * coefficients.minimum;
                var maximum = Math.Min(allowedMagnitude, requirementMagnitude * coefficients.maximum);
                if (maximum < minimum)
                {
                    throw new InvalidOperationException("Minimum-requirement coefficients do not overlap the configured measurement interval.");
                }

                return (minimum, maximum);
            }

            var boundedMinimum = Math.Min(allowedMagnitude, allowedMagnitude * coefficients.minimum);
            return (boundedMinimum, Math.Max(boundedMinimum, allowedMagnitude * coefficients.maximum));
        }

        private static double ResolveAllowedMagnitude(double standardValue, double lower, double upper, int direction)
        {
            var magnitude = direction > 0
                ? upper - standardValue
                : standardValue - lower;
            if (magnitude <= 0)
            {
                throw new InvalidOperationException("Measurement interval does not allow the selected deviation direction.");
            }

            return magnitude;
        }

        private double ResolveBaseMagnitude(MeasurementGenerationInput input, double minimumMagnitude, double maximumMagnitude, double allowedMagnitude)
        {
            if (input.AnchorError.HasValue)
            {
                var anchorMagnitude = Math.Abs(input.AnchorError.Value);
                var anchorSpread = allowedMagnitude * _configuration.ResultGroupMaximumFluctuationCoefficient;
                var anchorMinimum = Math.Max(minimumMagnitude, anchorMagnitude - anchorSpread);
                var anchorMaximum = Math.Min(maximumMagnitude, anchorMagnitude + anchorSpread);
                if (anchorMaximum >= anchorMinimum)
                {
                    return NextMagnitude(input.DistributionMode, anchorMinimum, anchorMaximum);
                }
            }

            return NextMagnitude(input.DistributionMode, minimumMagnitude, maximumMagnitude);
        }

    }
}
