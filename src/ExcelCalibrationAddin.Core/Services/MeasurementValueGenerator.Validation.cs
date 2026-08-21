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
        private static void ValidateInput(MeasurementGenerationInput input)
        {
            if (input == null)
            {
                throw new ArgumentNullException(nameof(input));
            }

            if (input.ValueCount <= 0)
            {
                throw new InvalidOperationException("Value count must be greater than zero.");
            }

            if (input.DecimalPlaces < 0 || input.DecimalPlaces > 15)
            {
                throw new InvalidOperationException("Decimal places must be between 0 and 15.");
            }

            if (input.DecimalPlacesByValue != null &&
                input.DecimalPlacesByValue.Any(value => value < 0 || value > 15))
            {
                throw new InvalidOperationException("Each decimal-place value must be between 0 and 15.");
            }

            if (double.IsNaN(input.StandardValue) || double.IsInfinity(input.StandardValue))
            {
                throw new InvalidOperationException("Standard value must be a finite number.");
            }

            if (double.IsNaN(input.Mpe) || double.IsInfinity(input.Mpe) || input.Mpe <= 0)
            {
                throw new InvalidOperationException("MPE must be greater than zero.");
            }

            if (input.NegativeTolerance.HasValue &&
                (double.IsNaN(input.NegativeTolerance.Value) || double.IsInfinity(input.NegativeTolerance.Value) || input.NegativeTolerance.Value < 0))
            {
                throw new InvalidOperationException("Negative tolerance must be a finite non-negative number.");
            }

            if (input.PositiveTolerance.HasValue &&
                (double.IsNaN(input.PositiveTolerance.Value) || double.IsInfinity(input.PositiveTolerance.Value) || input.PositiveTolerance.Value < 0))
            {
                throw new InvalidOperationException("Positive tolerance must be a finite non-negative number.");
            }

            if (input.ErrorType == ErrorType.Referenced &&
                (!input.ReferenceRange.HasValue ||
                 double.IsNaN(input.ReferenceRange.Value) ||
                 double.IsInfinity(input.ReferenceRange.Value) ||
                 input.ReferenceRange.Value <= 0))
            {
                throw new InvalidOperationException("Referenced error requires a reference range greater than zero.");
            }

            if (input.MeasurementLowerBound.HasValue &&
                (double.IsNaN(input.MeasurementLowerBound.Value) || double.IsInfinity(input.MeasurementLowerBound.Value)))
            {
                throw new InvalidOperationException("Measurement lower bound must be a finite number.");
            }

            if (input.MeasurementUpperBound.HasValue &&
                (double.IsNaN(input.MeasurementUpperBound.Value) || double.IsInfinity(input.MeasurementUpperBound.Value)))
            {
                throw new InvalidOperationException("Measurement upper bound must be a finite number.");
            }

            if (input.MeasurementLowerBound.HasValue &&
                input.MeasurementUpperBound.HasValue &&
                input.MeasurementLowerBound.Value > input.MeasurementUpperBound.Value)
            {
                throw new InvalidOperationException("Measurement lower bound cannot be greater than upper bound.");
            }
        }

        private (double lower, double upper) ResolveBounds(MeasurementGenerationInput input)
        {
            var bounds = ResolveToleranceBounds(input);
            if (input.MeasurementLowerBound.HasValue)
            {
                bounds.lower = Math.Max(bounds.lower, input.MeasurementLowerBound.Value);
            }

            if (input.MeasurementUpperBound.HasValue)
            {
                bounds.upper = Math.Min(bounds.upper, input.MeasurementUpperBound.Value);
            }

            if (bounds.lower > bounds.upper)
            {
                throw new InvalidOperationException("Measurement interval does not overlap with the allowed error range.");
            }

            return bounds;
        }

        private (double lower, double upper) ResolveToleranceBounds(MeasurementGenerationInput input)
        {
            var expansion = IsMinimumRequirement(input.RequirementOperator)
                ? _configuration.MinimumRequirementMaximumCoefficient
                : 1.0;
            if (input.NegativeTolerance.HasValue || input.PositiveTolerance.HasValue)
            {
                var negativeTolerance = input.NegativeTolerance ?? input.Mpe;
                var positiveTolerance = input.PositiveTolerance ?? input.Mpe;
                switch (input.ErrorType)
                {
                    case ErrorType.Relative:
                        return ResolveRelativeBounds(input.StandardValue, negativeTolerance * expansion, positiveTolerance * expansion);
                    case ErrorType.Referenced:
                        if (!input.ReferenceRange.HasValue)
                        {
                            throw new InvalidOperationException("Referenced error requires reference range.");
                        }

                        return (input.StandardValue - negativeTolerance * input.ReferenceRange.Value * expansion, input.StandardValue + positiveTolerance * input.ReferenceRange.Value * expansion);
                    default:
                        return (input.StandardValue - negativeTolerance * expansion, input.StandardValue + positiveTolerance * expansion);
                }
            }

            switch (input.ErrorType)
            {
                case ErrorType.Relative:
                    var ratio = input.Mpe;
                    return ResolveRelativeBounds(input.StandardValue, ratio * expansion, ratio * expansion);
                case ErrorType.Referenced:
                    if (!input.ReferenceRange.HasValue)
                    {
                        throw new InvalidOperationException("Referenced error requires reference range.");
                    }

                    var offset = input.Mpe * input.ReferenceRange.Value * expansion;
                    return (input.StandardValue - offset, input.StandardValue + offset);
                default:
                    return (input.StandardValue - input.Mpe * expansion, input.StandardValue + input.Mpe * expansion);
            }
        }

        private static bool IsMinimumRequirement(TechnicalRequirementOperator requirementOperator)
        {
            return requirementOperator == TechnicalRequirementOperator.GreaterThan ||
                requirementOperator == TechnicalRequirementOperator.GreaterThanOrEqual;
        }

        private static (double lower, double upper) ResolveRelativeBounds(
            double standardValue,
            double negativeTolerance,
            double positiveTolerance)
        {
            var negativeMagnitude = Math.Abs(standardValue) * negativeTolerance;
            var positiveMagnitude = Math.Abs(standardValue) * positiveTolerance;
            return (standardValue - negativeMagnitude, standardValue + positiveMagnitude);
        }

        private int ResolveDirection(MeasurementGenerationInput input)
        {
            return ResolveDirection(input, double.NegativeInfinity, double.PositiveInfinity);
        }

        private int ResolveDirection(MeasurementGenerationInput input, double lower, double upper)
        {
            if (input.ForcePositiveDirection && input.ForceNegativeDirection)
            {
                throw new InvalidOperationException("Generation direction flags conflict.");
            }

            if (lower >= input.StandardValue)
            {
                if (input.ForceNegativeDirection)
                {
                    throw new InvalidOperationException("Forced negative direction conflicts with the measurement interval.");
                }

                return 1;
            }

            if (upper <= input.StandardValue)
            {
                if (input.ForcePositiveDirection)
                {
                    throw new InvalidOperationException("Forced positive direction conflicts with the measurement interval.");
                }

                return -1;
            }

            if (input.ForcePositiveDirection)
            {
                return 1;
            }

            if (input.ForceNegativeDirection)
            {
                return -1;
            }

            return _random.Next(0, 2) == 0 ? -1 : 1;
        }

    }
}
