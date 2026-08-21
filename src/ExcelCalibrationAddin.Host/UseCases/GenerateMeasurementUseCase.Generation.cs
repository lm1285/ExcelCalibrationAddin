using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Generation;
using ExcelCalibrationAddin.Host.Recognition;
using ExcelCalibrationAddin.Host.Services;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public sealed partial class GenerateMeasurementUseCase
    {
        private static readonly Regex AverageRoundingFormulaRegex = new Regex(
            @"(?<function>ROUND(?:UP|DOWN)?)\s*\(.*[,;]\s*(?<digits>-?\d+)\s*\)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private RulePreview GeneratePreview(MeasurementRule rule, WorkbookSnapshot snapshot, MeasurementGenerationSession session)
        {
            var writableResolution = WritableCellResolver.Resolve(snapshot, rule.TargetRange);
            var writableCells = writableResolution.Cells;
            rule.WritableCells = CloneCellAddresses(writableCells);
            if (GenerationRuleValidator.IsAlarmRule(rule))
            {
                GenerationRuleValidator.ValidateAlarmRule(rule, writableCells.Count, writableResolution.FailureReason);
                return GenerateAlarmPreview(rule, writableCells, writableCells.Count);
            }

            if (GenerationRuleValidator.IsUpperLimitRule(rule))
            {
                GenerationRuleValidator.ValidateUpperLimitRule(rule, writableCells.Count, writableResolution.FailureReason);
                return GenerateUpperLimitPreview(rule, writableCells);
            }

            if (GenerationRuleValidator.IsRepeatabilityRule(rule))
            {
                GenerationRuleValidator.ValidateRepeatabilityRule(rule, writableCells.Count, writableResolution.FailureReason);
                var preview = GenerateRepeatabilityPreview(rule, writableCells, session);
                preview.WritableCells = CloneCellAddresses(writableCells);
                return preview;
            }

            ValidateStandardGenerationRule(rule, writableCells.Count, writableResolution.FailureReason);
            var generator = _generatorFactory(_generationConfiguration);
            return GenerateStandardValuePreview(rule, generator, snapshot, writableCells, session);
        }

        private MeasurementGenerationResult GenerateValues(
            MeasurementRule rule,
            MeasurementValueGenerator generator,
            double standardValue,
            int valueCount,
            int? forcedDirection,
            double? anchorError,
            IReadOnlyList<int> decimalPlacesByValue)
        {
            return generator.Generate(new MeasurementGenerationInput
            {
                StandardValue = standardValue,
                Mpe = rule.FixedMpe.Value,
                NegativeTolerance = rule.FixedNegativeTolerance,
                PositiveTolerance = rule.FixedPositiveTolerance,
                RequirementOperator = rule.RequirementOperator,
                ReferenceRange = rule.FixedReferenceRange,
                ErrorType = GenerationRuleValidator.ResolveGenerationErrorType(rule),
                DistributionMode = ResolveDistributionMode(_generationConfiguration),
                ValueCount = valueCount,
                DecimalPlaces = decimalPlacesByValue?.FirstOrDefault() ?? rule.FormatRule.DecimalPlaces ?? 2,
                DecimalPlacesByValue = decimalPlacesByValue?.ToList() ?? new List<int>(),
                ForcePositiveDirection = rule.PositiveDirectionOnly || forcedDirection > 0,
                ForceNegativeDirection = rule.NegativeDirectionOnly || forcedDirection < 0,
                AnchorError = anchorError,
                CoefficientOverride = rule.GenerationCoefficientOverride,
                MeasurementLowerBound = rule.MeasurementLowerBound,
                MeasurementUpperBound = rule.MeasurementUpperBound
            });
        }

        private static void ValidateStandardGenerationRule(MeasurementRule rule, int writableCellCount, string writableFailureReason = null)
        {
            var ruleName = GenerationRuleValidator.ResolveRuleName(rule);
            if (rule?.TargetRange == null)
            {
                throw new InvalidOperationException($"“{ruleName}”未设置测量值写入区域。");
            }

            if (!rule.FixedMpe.HasValue || rule.FixedMpe.Value <= 0)
            {
                throw new InvalidOperationException($"“{ruleName}”缺少有效的技术要求/允许误差。");
            }

            if (!rule.FixedStandardValue.HasValue &&
                GetManualStandardValuesByPoint(rule).Count == 0 &&
                !GenerationRuleValidator.HasValidRange(rule.StandardValueSource?.Range))
            {
                throw new InvalidOperationException($"“{ruleName}”缺少标准值。请检查标准值区域，或在侧边栏设置手动标准值。");
            }

            if (GenerationRuleValidator.ResolveGenerationErrorType(rule) == ErrorType.Referenced &&
                (!rule.FixedReferenceRange.HasValue || rule.FixedReferenceRange.Value <= 0))
            {
                throw new InvalidOperationException($"“{ruleName}”使用引用误差时必须提供有效量程。");
            }

            if (writableCellCount <= 0)
            {
                var message = $"“{ruleName}”的测量值写入区域无效。";
                throw new InvalidOperationException(string.IsNullOrWhiteSpace(writableFailureReason)
                    ? message
                    : $"{message}{Environment.NewLine}原因：{writableFailureReason}");
            }

            rule.GroupSize = writableCellCount;
        }

        private static double ResolveStandardValue(MeasurementRule rule)
        {
            if (rule.FixedStandardValue.HasValue)
            {
                return rule.FixedStandardValue.Value;
            }

            throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”缺少标准值。请检查标准值区域，或在侧边栏设置手动标准值。");
        }

        private static double CalculateRepresentativeError(MeasurementRule rule, double standardValue, IReadOnlyList<double> rawValues)
        {
            var representativeValue = rule.ErrorFormula != null && rule.ErrorFormula.ReferencesAverage
                ? ApplyAverageFormulaRounding(rawValues.Average(), rule.ErrorFormula.AverageFormula)
                : rawValues[0];
            return representativeValue - standardValue;
        }

        private static double ApplyAverageFormulaRounding(double value, string formula)
        {
            var match = AverageRoundingFormulaRegex.Match(formula ?? string.Empty);
            if (!match.Success ||
                !int.TryParse(match.Groups["digits"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var digits))
            {
                return value;
            }

            digits = Math.Max(-15, Math.Min(15, digits));
            var factor = Math.Pow(10, Math.Abs(digits));
            var scaled = digits >= 0 ? value * factor : value / factor;
            double rounded;
            switch (match.Groups["function"].Value.ToUpperInvariant())
            {
                case "ROUNDUP":
                    rounded = scaled >= 0 ? Math.Ceiling(scaled) : Math.Floor(scaled);
                    break;
                case "ROUNDDOWN":
                    rounded = Math.Truncate(scaled);
                    break;
                default:
                    rounded = Math.Round(scaled, MidpointRounding.AwayFromZero);
                    break;
            }

            return digits >= 0 ? rounded / factor : rounded * factor;
        }

        private static double CalculateFormulaError(MeasurementRule rule, double standardValue, IReadOnlyList<double> rawValues)
        {
            var error = CalculateRepresentativeError(rule, standardValue, rawValues);
            switch (rule?.ErrorFormula?.Scale)
            {
                case ErrorFormulaScale.RelativeToReferenceRange:
                    return rule.FixedReferenceRange.HasValue && Math.Abs(rule.FixedReferenceRange.Value) > 1e-12
                        ? ScaleFormulaRatio(error / rule.FixedReferenceRange.Value, rule.ErrorFormula)
                        : 0;
                case ErrorFormulaScale.RelativeToStandardValue:
                    return Math.Abs(standardValue) > 1e-12
                        ? ScaleFormulaRatio(error / standardValue, rule.ErrorFormula)
                        : 0;
                default:
                    return error;
            }
        }

        private static double ScaleFormulaRatio(double ratio, ErrorFormulaInfo formula)
        {
            return formula?.FormulaMultipliesBy100 == true ? ratio * 100.0 : ratio;
        }

        private static double? ResolveAnchorError(
            MeasurementRule rule,
            double standardValue,
            double? anchorErrorRatio,
            double? anchorErrorMagnitude)
        {
            if (rule == null)
            {
                return null;
            }

            var errorType = GenerationRuleValidator.ResolveGenerationErrorType(rule);
            if (errorType == ErrorType.Relative && anchorErrorRatio.HasValue)
            {
                return anchorErrorRatio.Value * standardValue;
            }

            if (errorType == ErrorType.Referenced &&
                anchorErrorMagnitude.HasValue &&
                rule.FixedReferenceRange.HasValue &&
                rule.FixedReferenceRange.Value > 0)
            {
                return anchorErrorMagnitude.Value;
            }

            return anchorErrorMagnitude;
        }

        private static void ValidateGeneratedError(MeasurementRule rule, double standardValue, double error)
        {
            var allowedMagnitude = ResolveAllowedFormulaErrorMagnitude(rule, standardValue);
            if (allowedMagnitude.HasValue && IsRequirementSatisfied(rule.RequirementOperator, error, allowedMagnitude.Value))
            {
                return;
            }

            if (allowedMagnitude.HasValue && rule.RequirementOperator != TechnicalRequirementOperator.None)
            {
                throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”生成值不满足技术要求运算符。");
            }

            var bounds = ResolveFormulaErrorBounds(rule, standardValue);
            if (error < bounds.lower || error > bounds.upper)
            {
            throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”生成值回代误差公式后超出技术要求。");
            }
        }

        private static bool IsRequirementSatisfied(
            TechnicalRequirementOperator requirementOperator,
            double error,
            double limit)
        {
            var magnitude = Math.Abs(error);
            var tolerance = Math.Max(1e-12, Math.Abs(limit) * 1e-12);
            switch (requirementOperator)
            {
                case TechnicalRequirementOperator.LessThan:
                    return magnitude < limit || (magnitude > limit && magnitude - limit <= tolerance);
                case TechnicalRequirementOperator.LessThanOrEqual:
                case TechnicalRequirementOperator.PlusMinus:
                    return magnitude <= limit + tolerance;
                case TechnicalRequirementOperator.GreaterThan:
                    return magnitude > limit || (magnitude < limit && limit - magnitude <= tolerance);
                case TechnicalRequirementOperator.GreaterThanOrEqual:
                    return magnitude >= limit - tolerance;
                default:
                    return false;
            }
        }

        private void ValidateConfiguredErrorUsage(MeasurementRule rule, double standardValue, double error)
        {
            var usage = ResolveConfiguredErrorUsage(rule, standardValue, error);
            if (!usage.HasValue)
            {
                return;
            }

            var coefficients = ResolveRangeCoefficients(rule, usage.Value < 0 ? -1 : 1);
            var magnitude = Math.Abs(usage.Value);
            if (magnitude + 1e-12 < coefficients.minimum || magnitude - 1e-12 > coefficients.maximum)
            {
            throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”生成误差未落在配置的误差占用比例内。");
            }
        }

        private double? ResolveConfiguredErrorUsage(MeasurementRule rule, double standardValue, double error)
        {
            if (rule == null)
            {
                return null;
            }

            var allowedMagnitude = ResolveAllowedFormulaErrorMagnitude(rule, standardValue);
            if (!allowedMagnitude.HasValue || allowedMagnitude.Value <= 0)
            {
                return null;
            }

            return error / allowedMagnitude.Value;
        }

        private static double? ResolveAllowedFormulaErrorMagnitude(MeasurementRule rule, double standardValue)
        {
            if (rule == null)
            {
                return null;
            }

            var negativeTolerance = rule.FixedNegativeTolerance ?? rule.FixedMpe.GetValueOrDefault();
            var positiveTolerance = rule.FixedPositiveTolerance ?? rule.FixedMpe.GetValueOrDefault();
            var tolerance = Math.Max(Math.Abs(negativeTolerance), Math.Abs(positiveTolerance));
            if (tolerance <= 0)
            {
                return null;
            }

            switch (rule.ErrorFormula?.Scale)
            {
                case ErrorFormulaScale.RelativeToStandardValue:
                    return ScaleFormulaRatio(tolerance, rule.ErrorFormula);
                case ErrorFormulaScale.RelativeToReferenceRange:
                    return ScaleFormulaRatio(tolerance, rule.ErrorFormula);
                default:
                    switch (rule.ErrorType)
                    {
                        case ErrorType.Relative:
                            return Math.Abs(standardValue) * tolerance;
                        case ErrorType.Referenced:
                            return rule.FixedReferenceRange.HasValue
                                ? Math.Abs(rule.FixedReferenceRange.Value) * tolerance
                                : (double?)null;
                        default:
                            return tolerance;
                    }
            }
        }

        private (double minimum, double maximum) ResolveRangeCoefficients(MeasurementRule rule, int direction)
        {
            if (rule != null &&
                (rule.RequirementOperator == TechnicalRequirementOperator.GreaterThan ||
                 rule.RequirementOperator == TechnicalRequirementOperator.GreaterThanOrEqual))
            {
                return (
                    _generationConfiguration.MinimumRequirementMinimumCoefficient,
                    _generationConfiguration.MinimumRequirementMaximumCoefficient);
            }

            var coefficientOverride = rule?.GenerationCoefficientOverride;
            if (_generationConfiguration?.UseIndependentDeviationControl != true)
            {
                return (
                    coefficientOverride?.PositiveMinimumCoefficient ??
                    coefficientOverride?.NegativeMinimumCoefficient ??
                    _generationConfiguration.UnifiedErrorMinimumCoefficient,
                    coefficientOverride?.PositiveMaximumCoefficient ??
                    coefficientOverride?.NegativeMaximumCoefficient ??
                    _generationConfiguration.UnifiedErrorMaximumCoefficient);
            }

            if (direction > 0)
            {
                return (
                    coefficientOverride?.PositiveMinimumCoefficient ?? _generationConfiguration.PositiveErrorMinimumCoefficient,
                    coefficientOverride?.PositiveMaximumCoefficient ?? _generationConfiguration.PositiveErrorMaximumCoefficient);
            }

            return (
                coefficientOverride?.NegativeMinimumCoefficient ?? _generationConfiguration.NegativeErrorMinimumCoefficient,
                coefficientOverride?.NegativeMaximumCoefficient ?? _generationConfiguration.NegativeErrorMaximumCoefficient);
        }

        private static (double lower, double upper) ResolveErrorBounds(MeasurementRule rule, double standardValue)
        {
            var negativeTolerance = rule.FixedNegativeTolerance ?? rule.FixedMpe.GetValueOrDefault();
            var positiveTolerance = rule.FixedPositiveTolerance ?? rule.FixedMpe.GetValueOrDefault();
            switch (rule.ErrorType)
            {
                case ErrorType.Relative:
                    return (-Math.Abs(standardValue * negativeTolerance), Math.Abs(standardValue * positiveTolerance));
                case ErrorType.Referenced:
                    var referenceRange = rule.FixedReferenceRange.GetValueOrDefault();
                    return (-Math.Abs(referenceRange * negativeTolerance), Math.Abs(referenceRange * positiveTolerance));
                default:
                    return (-Math.Abs(negativeTolerance), Math.Abs(positiveTolerance));
            }
        }

        private static (double lower, double upper) ResolveFormulaErrorBounds(MeasurementRule rule, double standardValue)
        {
            var negativeTolerance = rule.FixedNegativeTolerance ?? rule.FixedMpe.GetValueOrDefault();
            var positiveTolerance = rule.FixedPositiveTolerance ?? rule.FixedMpe.GetValueOrDefault();
            switch (rule.ErrorFormula?.Scale)
            {
                case ErrorFormulaScale.RelativeToStandardValue:
                case ErrorFormulaScale.RelativeToReferenceRange:
                    return (
                        -Math.Abs(ScaleFormulaRatio(negativeTolerance, rule.ErrorFormula)),
                        Math.Abs(ScaleFormulaRatio(positiveTolerance, rule.ErrorFormula)));
                default:
                    return ResolveErrorBounds(rule, standardValue);
            }
        }

    }
}
