using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
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
        private static IReadOnlyList<MeasurementRule> NormalizeRules(IEnumerable<MeasurementRule> rules, WorkbookSnapshot snapshot)
        {
            var normalizedRules = (rules ?? Enumerable.Empty<MeasurementRule>())
                .Where(rule => rule != null)
                .Where(rule => rule.IsEnabled)
                .ToList();

            foreach (var rule in normalizedRules)
            {
                NormalizeMpeScaleForFormula(rule);
            }

            var enabledRules = normalizedRules;
            normalizedRules = enabledRules
                .Where(rule => IsGenerationReady(rule, snapshot))
                .ToList();

            if (normalizedRules.Count == 0)
            {
                if (enabledRules.Any(rule =>
                    rule != null &&
                    rule.IsEnabled &&
                    GenerationRuleValidator.IsAlarmRule(rule) &&
                    GenerationRuleValidator.HasValidRange(rule.TargetRange) &&
                    !rule.FixedStandardValue.HasValue))
                {
                    throw new InvalidOperationException("请先在功能区的“报警值输入”中输入具体数值。");
                }

                throw new InvalidOperationException("没有可生成的校准规则。请先识别模板并确认测量值区域、标准值和允许误差。");
            }

            return normalizedRules;
        }

        private static bool IsGenerationReady(MeasurementRule rule, WorkbookSnapshot snapshot)
        {
            if (GenerationRuleValidator.IsNonNumericRule(rule))
            {
                return false;
            }

            if (GenerationRuleValidator.IsAlarmRule(rule))
            {
                return rule != null &&
                    GenerationRuleValidator.HasValidRange(rule.TargetRange) &&
                    rule.FixedStandardValue.HasValue;
            }

            if (GenerationRuleValidator.IsUpperLimitRule(rule))
            {
                return rule != null &&
                    GenerationRuleValidator.HasValidRange(rule.TargetRange) &&
                    rule.FixedMpe.HasValue &&
                    rule.FixedMpe.Value > 0;
            }

            if (GenerationRuleValidator.IsRepeatabilityRule(rule))
            {
                return rule != null &&
                    GenerationRuleValidator.HasValidRange(rule.TargetRange) &&
                    rule.FixedStandardValue.HasValue &&
                    rule.FixedMpe.HasValue &&
                    rule.FixedMpe.Value > 0;
            }

            return rule != null &&
                GenerationRuleValidator.HasValidRange(rule.TargetRange) &&
                HasUsableStandardValue(rule, snapshot) &&
                rule.FixedMpe.HasValue &&
                rule.FixedMpe.Value > 0 &&
                (GenerationRuleValidator.ResolveGenerationErrorType(rule) != ErrorType.Referenced ||
                 (rule.FixedReferenceRange.HasValue && rule.FixedReferenceRange.Value > 0));
        }

        public IReadOnlyList<RulePreview> Preview(IEnumerable<MeasurementRule> rules)
        {
            var ruleList = (rules ?? Enumerable.Empty<MeasurementRule>()).ToList();
            var snapshot = CaptureRuleRanges(ruleList);
            var normalizedRules = NormalizeRules(ResolveRules(ruleList, snapshot), snapshot);
            return PreviewNormalized(normalizedRules, snapshot);
        }

        public IReadOnlyList<RulePreview> PreviewResolved(IEnumerable<MeasurementRule> rules)
        {
            var ruleList = (rules ?? Enumerable.Empty<MeasurementRule>()).ToList();
            var snapshot = CaptureRuleRanges(ruleList);
            var normalizedRules = NormalizeRules(ResolveRules(ruleList, snapshot), snapshot);
            return PreviewNormalized(normalizedRules, snapshot);
        }

        private WorkbookSnapshot CaptureRuleRanges(IReadOnlyList<MeasurementRule> rules)
        {
            if (_snapshotProvider == null)
            {
                return null;
            }

            var ranges = (rules ?? Array.Empty<MeasurementRule>())
                .Where(rule => rule != null)
                .SelectMany(rule => new[]
                {
                    rule.TargetRange,
                    rule.StandardValueSource?.Range,
                    rule.AverageSource?.Range,
                    rule.ErrorSource?.Range,
                    rule.MpeSource?.Range,
                    rule.RangeSource?.Range,
                    rule.UncertaintySource?.Range,
                    rule.ResultSource?.Range
                })
                .Where(GenerationRuleValidator.HasValidRange)
                .ToList();

            return ranges.Count == 0 ? _snapshotProvider.Capture() : _snapshotProvider.Capture(ranges);
        }

        public IReadOnlyList<RulePreview> PreviewPreResolved(IEnumerable<MeasurementRule> rules)
        {
            var ruleList = (rules ?? Enumerable.Empty<MeasurementRule>()).ToList();
            var snapshot = CaptureRuleRanges(ruleList);
            var resolvedFormulaScales = ruleList
                .Select(rule => rule?.ErrorFormula?.Scale)
                .ToList();
            _structureAnalyzer.Apply(snapshot, ruleList);
            for (var index = 0; index < ruleList.Count; index++)
            {
                var scale = resolvedFormulaScales[index];
                var rule = ruleList[index];
                var formula = rule?.ErrorFormula;
                if (scale != ErrorFormulaScale.RelativeToReferenceRange ||
                    rule?.ErrorType != ErrorType.Referenced ||
                    formula?.Scale != ErrorFormulaScale.RelativeToStandardValue)
                {
                    continue;
                }

                formula.Scale = ErrorFormulaScale.RelativeToReferenceRange;
                formula.FormulaDividesByReferenceRange = true;
            }
            var normalizedRules = NormalizeRules(ruleList, snapshot);
            return PreviewPreResolvedNormalized(normalizedRules, snapshot);
        }

        private IReadOnlyList<RulePreview> PreviewNormalized(IReadOnlyList<MeasurementRule> normalizedRules, WorkbookSnapshot snapshot)
        {
            var previews = new List<RulePreview>(normalizedRules.Count);
            var session = new MeasurementGenerationSession(snapshot);

            foreach (var rule in normalizedRules)
            {
                var preview = GeneratePreview(rule, snapshot, session);
                previews.Add(preview);
            }

            return previews;
        }

        private IReadOnlyList<RulePreview> PreviewPreResolvedNormalized(IReadOnlyList<MeasurementRule> normalizedRules, WorkbookSnapshot snapshot)
        {
            var previews = new List<RulePreview>(normalizedRules.Count);
            var session = new MeasurementGenerationSession(snapshot);

            foreach (var rule in normalizedRules)
            {
                var preview = GeneratePreResolvedPreview(rule, snapshot, session);
                previews.Add(preview);
            }

            return previews;
        }

        private RulePreview GeneratePreResolvedPreview(MeasurementRule rule, WorkbookSnapshot snapshot, MeasurementGenerationSession session)
        {
            var writableResolution = WritableCellResolver.Resolve(snapshot, rule?.TargetRange);
            var resolvedWritableCells = writableResolution.Cells;
            var writableCells = resolvedWritableCells.Count > 0
                ? resolvedWritableCells
                : HasWritableCells(rule)
                    ? CloneCellAddresses(rule.WritableCells)
                    : null;
            var writableCellCount = writableCells?.Count > 0
                ? writableCells.Count
                : rule?.GroupSize > 0
                ? rule.GroupSize
                : WritableCellResolver.CountRangeCells(rule?.TargetRange);
            writableCells = writableCells?.Count > 0
                ? writableCells
                : BuildContiguousWritableCells(rule?.TargetRange, writableCellCount);
            var writableFailureReason = writableCells?.Count > 0
                ? null
                : writableResolution.FailureReason;
            if (GenerationRuleValidator.IsAlarmRule(rule))
            {
                GenerationRuleValidator.ValidateAlarmRule(rule, writableCellCount, writableFailureReason);
                return GenerateAlarmPreview(rule, writableCells, writableCellCount);
            }

            if (GenerationRuleValidator.IsUpperLimitRule(rule))
            {
                GenerationRuleValidator.ValidateUpperLimitRule(rule, writableCellCount, writableFailureReason);
                return GenerateUpperLimitPreview(rule, writableCells);
            }

            if (GenerationRuleValidator.IsRepeatabilityRule(rule))
            {
                GenerationRuleValidator.ValidateRepeatabilityRule(rule, writableCellCount, writableFailureReason);
                var preview = GenerateRepeatabilityPreview(rule, writableCells, session);
                preview.WritableCells = writableCells;
                return preview;
            }

            ValidateStandardGenerationRule(rule, writableCellCount, writableFailureReason);
            var generator = _generatorFactory(_generationConfiguration);
            return GenerateStandardValuePreview(rule, generator, snapshot, writableCells, session);
        }

        private static RulePreview GenerateAlarmPreview(
            MeasurementRule rule,
            IReadOnlyList<CellAddress> writableCells,
            int writableCellCount)
        {
            var alarmValue = rule.FixedStandardValue.Value;
            var displayValue = alarmValue.ToString("G15", CultureInfo.InvariantCulture);
            return new RulePreview
            {
                Rule = rule,
                TargetRange = rule.TargetRange,
                DisplayValues = Enumerable.Repeat(displayValue, writableCellCount).ToList(),
                RawValues = Enumerable.Repeat(alarmValue, writableCellCount).ToList(),
                WritableCells = writableCells ?? new List<CellAddress>()
            };
        }

        private RulePreview GenerateRepeatabilityPreview(
            MeasurementRule rule,
            IReadOnlyList<CellAddress> writableCells,
            MeasurementGenerationSession session)
        {
            var standardValue = ResolveStandardValue(rule);
            var toleranceRatio = MeasurementSeriesGenerator.ResolveRepeatabilityTolerance(rule.FixedMpe.Value);
            var decimalPlacesByValue = ResolveDecimalPlaces(session, rule, writableCells);
            var decimalPlaces = decimalPlacesByValue.DefaultIfEmpty(rule.FormatRule?.DecimalPlaces ?? 2).Min();
            var errorDecimalPlaces = ResolveErrorDecimalPlaces(session, rule);
            var minimumVisibleSpread = ResolveMinimumVisibleRepeatabilitySpread(
                rule,
                standardValue,
                errorDecimalPlaces);
            var crossItemKey = session.BuildCrossItemKey(rule, standardValue, ResolveMeasurementUnit(session, rule));
            var sharedProfile = session.FindCrossItem(crossItemKey);
            var centerValue = standardValue + (sharedProfile?.RepresentativeError ?? 0);
            if (rule.MeasurementLowerBound.HasValue)
            {
                centerValue = Math.Max(centerValue, rule.MeasurementLowerBound.Value);
            }

            if (rule.MeasurementUpperBound.HasValue)
            {
                centerValue = Math.Min(centerValue, rule.MeasurementUpperBound.Value);
            }

            var rawValues = _seriesGenerator.GenerateRepeatabilityValues(
                rule,
                standardValue,
                centerValue,
                toleranceRatio,
                writableCells.Count,
                decimalPlaces,
                _generationConfiguration,
                minimumVisibleSpread);
            var result = BuildGenerationResult(rawValues, decimalPlacesByValue, sharedProfile?.Direction ?? 0);
            var roundedValues = result.RawValues
                .Select((value, index) => Math.Round(value, decimalPlacesByValue[index]))
                .ToList();
            if (roundedValues.Distinct().Count() < 2)
            {
                throw new RepeatabilityVerificationException(
                    $"“{GenerationRuleValidator.ResolveRuleName(rule)}”在当前小数分辨力下无法生成非零重复性误差。");
            }

            var displayedRepeatability = Math.Round(
                CalculateRepeatabilityMetric(rule, standardValue, roundedValues),
                Math.Max(0, Math.Min(15, errorDecimalPlaces)));
            if (Math.Abs(displayedRepeatability) <= 1e-12)
            {
                throw new RepeatabilityVerificationException(
                    $"“{GenerationRuleValidator.ResolveRuleName(rule)}”在当前误差列精度下无法生成非零重复性误差。");
            }

            return new RulePreview
            {
                Rule = rule,
                TargetRange = rule.TargetRange,
                DisplayValues = result.DisplayValues,
                RawValues = result.RawValues
            };
        }

        private static double ResolveMinimumVisibleRepeatabilitySpread(
            MeasurementRule rule,
            double standardValue,
            int errorDecimalPlaces)
        {
            var errorResolution = Math.Pow(10, -Math.Max(0, Math.Min(15, errorDecimalPlaces)));
            var minimumDisplayedMetric = errorResolution * 0.55;
            var formula = rule?.ErrorFormula;
            double requiredStandardDeviation;
            switch (formula?.Scale)
            {
                case ErrorFormulaScale.RelativeToReferenceRange:
                    requiredStandardDeviation = minimumDisplayedMetric * Math.Abs(rule.FixedReferenceRange.GetValueOrDefault());
                    if (formula.FormulaMultipliesBy100)
                    {
                        requiredStandardDeviation /= 100d;
                    }
                    break;
                case ErrorFormulaScale.RelativeToStandardValue:
                    requiredStandardDeviation = minimumDisplayedMetric * Math.Abs(standardValue);
                    if (formula.FormulaMultipliesBy100)
                    {
                        requiredStandardDeviation /= 100d;
                    }
                    break;
                default:
                    if (formula?.HasFormula == true)
                    {
                        requiredStandardDeviation = minimumDisplayedMetric;
                    }
                    else
                    {
                        requiredStandardDeviation = minimumDisplayedMetric * Math.Abs(standardValue) / 100d;
                    }
                    break;
            }

            // Evenly distributed samples have a standard deviation of at least about
            // 28% of their full span. This margin keeps the displayed result away
            // from the rounding midpoint.
            return requiredStandardDeviation / 0.28d;
        }

        private static double CalculateRepeatabilityMetric(
            MeasurementRule rule,
            double standardValue,
            IReadOnlyList<double> values)
        {
            if (values == null || values.Count <= 1)
            {
                return 0;
            }

            var average = values.Average();
            var standardDeviation = Math.Sqrt(
                values.Sum(value => Math.Pow(value - average, 2)) / (values.Count - 1));
            var formula = rule?.ErrorFormula;
            switch (formula?.Scale)
            {
                case ErrorFormulaScale.RelativeToReferenceRange:
                    var referenceRange = Math.Abs(rule.FixedReferenceRange.GetValueOrDefault());
                    return referenceRange <= 1e-12
                        ? 0
                        : ScaleFormulaRatio(standardDeviation / referenceRange, formula);
                case ErrorFormulaScale.RelativeToStandardValue:
                    var absoluteStandardValue = Math.Abs(standardValue);
                    return absoluteStandardValue <= 1e-12
                        ? 0
                        : ScaleFormulaRatio(standardDeviation / absoluteStandardValue, formula);
                default:
                    return formula?.HasFormula == true
                        ? standardDeviation
                        : Math.Abs(average) <= 1e-12
                            ? standardDeviation
                            : standardDeviation / Math.Abs(average) * 100d;
            }
        }

        private RulePreview GenerateUpperLimitPreview(
            MeasurementRule rule,
            IReadOnlyList<CellAddress> writableCells)
        {
            var upperLimit = MeasurementSeriesGenerator.ResolveUpperLimit(rule);
            var decimalPlaces = rule.FormatRule?.DecimalPlaces ?? 2;
            var manualStandardValues = GetManualStandardValuesByPoint(rule);
            var orderedCells = (writableCells ?? Array.Empty<CellAddress>())
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();
            var rawValues = new List<double>();
            var generatedWritableCells = new List<CellAddress>();
            var isPlusMinusRequirement = IsPlusMinusRequirement(rule);

            if (!isPlusMinusRequirement && manualStandardValues.Count > 0)
            {
                var targetRows = orderedCells
                    .Select(cell => cell.Row)
                    .Distinct()
                    .OrderBy(row => row)
                    .ToList();
                if (targetRows.Count == 1 && manualStandardValues.Count > 1)
                {
                    var lowerBound = manualStandardValues.Values.Min();
                    var upperBound = manualStandardValues.Values.Max();
                    var centerValue = (lowerBound + upperBound) / 2d;
                    rawValues.AddRange(_seriesGenerator.GenerateResponseTimeValues(
                        centerValue,
                        upperLimit,
                        orderedCells.Count,
                        decimalPlaces,
                        _generationConfiguration,
                        lowerBound,
                        upperBound));
                    generatedWritableCells.AddRange(orderedCells);
                }
                else
                {
                    foreach (var manualStandardValue in manualStandardValues.OrderBy(item => item.Key))
                    {
                        if (manualStandardValue.Key > targetRows.Count)
                        {
                            continue;
                        }

                        var row = targetRows[manualStandardValue.Key - 1];
                        var rowCells = orderedCells.Where(cell => cell.Row == row).ToList();
                        rawValues.AddRange(_seriesGenerator.GenerateResponseTimeValues(
                            manualStandardValue.Value,
                            upperLimit,
                            rowCells.Count,
                            decimalPlaces,
                            _generationConfiguration,
                            rule.MeasurementLowerBound,
                            rule.MeasurementUpperBound));
                        generatedWritableCells.AddRange(rowCells);
                    }
                }
            }
            else if (!isPlusMinusRequirement)
            {
                var centerValue = rule.FixedStandardValue ??
                    _seriesGenerator.GenerateUpperLimitValues(
                        rule,
                        upperLimit,
                        1,
                        decimalPlaces,
                        _generationConfiguration)[0];
                rawValues.AddRange(_seriesGenerator.GenerateResponseTimeValues(
                    centerValue,
                    upperLimit,
                    orderedCells.Count,
                    decimalPlaces,
                    _generationConfiguration,
                    rule.MeasurementLowerBound,
                    rule.MeasurementUpperBound));
                generatedWritableCells.AddRange(orderedCells);
            }
            else
            {
                rawValues.AddRange(_seriesGenerator.GenerateUpperLimitValues(
                    rule,
                    upperLimit,
                    orderedCells.Count,
                    decimalPlaces,
                    _generationConfiguration));
                generatedWritableCells.AddRange(orderedCells);
            }

            if (rawValues.Count == 0)
            {
                throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”所有手动标准值位置均超出测量区域。");
            }

            return new RulePreview
            {
                Rule = rule,
                TargetRange = rule.TargetRange,
                DisplayValues = MeasurementSeriesGenerator.FormatValues(rawValues, decimalPlaces),
                RawValues = rawValues,
                WritableCells = CloneCellAddresses(generatedWritableCells)
            };
        }

        private static bool IsPlusMinusRequirement(MeasurementRule rule)
        {
            if (rule?.RequirementOperator == TechnicalRequirementOperator.PlusMinus)
            {
                return true;
            }

            return MpeValuePatternCodec.Parse(rule?.MpeSource?.ValuePattern)?.Operator ==
                TechnicalRequirementOperator.PlusMinus;
        }

        private IReadOnlyList<MeasurementRule> ResolveRules(IEnumerable<MeasurementRule> rules, WorkbookSnapshot snapshot)
        {
            var ruleList = (rules ?? Enumerable.Empty<MeasurementRule>())
                .Where(rule => rule != null)
                .Where(rule => rule.IsEnabled)
                .ToList();
            if (ruleList.Count == 0 || snapshot == null || _parameterResolver == null)
            {
                return ruleList;
            }

            var resolvedRules = _parameterResolver.Apply(snapshot, ruleList);
            return resolvedRules;
        }

        private RulePreview GenerateStandardValuePreview(
            MeasurementRule rule,
            MeasurementValueGenerator generator,
            WorkbookSnapshot snapshot,
            IReadOnlyList<CellAddress> writableCells,
            MeasurementGenerationSession session)
        {
            var targetRows = writableCells
                .Select(cell => cell.Row)
                .Distinct()
                .OrderBy(row => row)
                .ToList();
            var manualStandardValuesByPoint = GetManualStandardValuesByPoint(rule);
            var standardValuesByRow = HasManualStandardMode(rule)
                ? new Dictionary<int, double>()
                : ResolveStandardValuesByRow(snapshot, rule.StandardValueSource?.Range, targetRows);
            if (HasManualStandardMode(rule))
            {
                if (manualStandardValuesByPoint.Keys.All(pointIndex => pointIndex > targetRows.Count))
                {
                    throw new InvalidOperationException($"“{GenerationRuleValidator.ResolveRuleName(rule)}”所有手动标准值位置均超出测量区域。");
                }

                standardValuesByRow = manualStandardValuesByPoint
                    .Where(item => item.Key <= targetRows.Count)
                    .ToDictionary(item => targetRows[item.Key - 1], item => item.Value);
            }
            var displayValues = new List<string>();
            var rawValues = new List<double>();
            var generatedWritableCells = new List<CellAddress>();
            var trendKey = session.BuildTrendKey(rule, ResolveMeasurementUnit(session, rule));

            if (HasRowMappedStandardValues(standardValuesByRow, targetRows))
            {
                int? sharedDirection = null;
                double? anchorErrorRatio = null;
                double? anchorErrorMagnitude = null;
                var trendErrors = LoadTrendErrors(session, trendKey, rule);
                for (var rowIndex = 0; rowIndex < targetRows.Count; rowIndex++)
                {
                    var row = targetRows[rowIndex];
                    var valueCount = writableCells.Count(cell => cell.Row == row);
                    if (valueCount <= 0)
                    {
                        continue;
                    }

                    if (!TryResolveStandardValueForRow(rule, standardValuesByRow, row, out var standardValue))
                    {
                        continue;
                    }

                    var measurementCells = writableCells.Where(cell => cell.Row == row).OrderBy(cell => cell.Column).ToList();
                    var pointRule = ResolvePointRule(snapshot, rule, row, standardValue, measurementCells);
                    var anchorError = ResolveAnchorError(pointRule, standardValue, anchorErrorRatio, anchorErrorMagnitude);
                    var point = GeneratePointValues(
                        pointRule,
                        generator,
                        snapshot,
                        measurementCells,
                        standardValue,
                        _generationConfiguration?.UseSameDeviationDirection == true ? sharedDirection : null,
                        anchorError,
                        session,
                        trendErrors,
                        trendKey);
                    var result = point.Result;
                    var representativeError = point.RepresentativeError;
                    if (rowIndex == 0)
                    {
                        sharedDirection = result.Direction == 0
                            ? (representativeError < 0 ? -1 : 1)
                            : result.Direction;
                        anchorErrorMagnitude = Math.Abs(representativeError);
                        anchorErrorRatio = Math.Abs(standardValue) > 1e-12
                            ? representativeError / standardValue
                            : (double?)null;
                    }

                    displayValues.AddRange(result.DisplayValues);
                    rawValues.AddRange(result.RawValues);
                    generatedWritableCells.AddRange(writableCells.Where(cell => cell.Row == row));
                }
            }
            else
            {
                var standardValue = ResolveStandardValue(rule);
                var orderedCells = writableCells.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToList();
                var point = GeneratePointValues(
                    rule,
                    generator,
                    snapshot,
                    orderedCells,
                    standardValue,
                    null,
                    null,
                    session,
                    LoadTrendErrors(session, trendKey, rule),
                    trendKey);
                var result = point.Result;
                displayValues.AddRange(result.DisplayValues);
                rawValues.AddRange(result.RawValues);
                generatedWritableCells.AddRange(writableCells);
            }

            if (displayValues.Count == 0)
            {
                return new RulePreview
                {
                    Rule = rule,
                    TargetRange = rule.TargetRange,
                    DisplayValues = new List<string>(),
                    RawValues = new List<double>(),
                    WritableCells = new List<CellAddress>(),
                    WarningMessages = new[]
                    {
                        $"“{GenerationRuleValidator.ResolveRuleName(rule)}”所有标准值行均为空，已跳过生成。"
                    }
                };
            }

            return new RulePreview
            {
                Rule = rule,
                TargetRange = rule.TargetRange,
                DisplayValues = displayValues,
                RawValues = rawValues,
                WritableCells = CloneCellAddresses(generatedWritableCells),
                WarningMessages = BuildGenerationWarnings(rule)
            };
        }

        private GeneratedPoint GeneratePointValues(
            MeasurementRule rule,
            MeasurementValueGenerator generator,
            WorkbookSnapshot snapshot,
            IReadOnlyList<CellAddress> measurementCells,
            double standardValue,
            int? forcedDirection,
            double? anchorError,
            MeasurementGenerationSession session,
            ICollection<TrendErrorSample> trendErrors,
            string trendKey)
        {
            ValidateStandardGenerationRule(rule, measurementCells?.Count ?? 0);
            var decimalPlaces = ResolveDecimalPlaces(session, rule, measurementCells);
            var errorDecimalPlaces = ResolveErrorDecimalPlaces(session, rule);
            var resolvedUnit = ResolveMeasurementUnit(session, rule);
            var sharedKey = session.BuildKey(rule, standardValue, resolvedUnit) +
                "|resolution:" + string.Join(",", decimalPlaces);
            var crossItemKey = session.BuildCrossItemKey(rule, standardValue, resolvedUnit);
            var sharedProfile = session.Find(sharedKey);
            if (sharedProfile != null)
            {
                EnsureSharedDirectionIsCompatible(rule, sharedProfile.Direction);
                try
                {
                    if (sharedProfile.RawValues.Count >= measurementCells.Count)
                    {
                        var reusedValues = sharedProfile.RawValues.Take(measurementCells.Count).ToList();
                        EnsureValuesMatchResolution(rule, reusedValues, decimalPlaces);
                        EnsureValuesMatchIndependentInterval(rule, reusedValues);
                        var reusedResult = BuildGenerationResult(reusedValues, decimalPlaces, sharedProfile.Direction);
                        var reusedError = CalculateRepresentativeError(rule, standardValue, reusedResult.RawValues);
                        var reusedTrendError = CalculateRoundedFormulaError(
                            rule,
                            standardValue,
                            reusedResult.RawValues,
                            errorDecimalPlaces);
                        if (Math.Abs(reusedTrendError) <= 1e-12)
                        {
                            throw new InvalidOperationException(
                                "Generated calibration error is zero after applying the error-field precision.");
                        }
                        ValidateGeneratedError(rule, standardValue, reusedTrendError);
                        ValidateConfiguredErrorUsage(rule, standardValue, reusedTrendError);
                        if (TryAddTrendError(rule, standardValue, reusedTrendError, trendErrors))
                        {
                            session.AddTrendError(trendKey, standardValue, reusedTrendError);
                            return new GeneratedPoint(reusedResult, reusedError);
                        }

                        sharedProfile = null;
                    }
                }
                catch (InvalidOperationException)
                {
                    sharedProfile = null;
                }
            }

            var effectiveDirection = forcedDirection ?? sharedProfile?.Direction;
            var effectiveAnchor = sharedProfile?.RepresentativeError ?? anchorError;
            const int maximumAttempts = 32;
            var technicalRequirementRejectCount = 0;
            var configuredIntervalRejectCount = 0;
            var trendRejectCount = 0;
            var duplicateRejectCount = 0;
            double? minimumRejectedUsage = null;
            double? maximumRejectedUsage = null;
            double? minimumRejectedRepresentativeError = null;
            double? maximumRejectedRepresentativeError = null;
            for (var attempt = 0; attempt < maximumAttempts; attempt++)
            {
                var result = GenerateValues(
                    rule,
                    generator,
                    standardValue,
                    measurementCells.Count,
                    effectiveDirection,
                    effectiveAnchor,
                    decimalPlaces);
                if (sharedProfile != null && sharedProfile.RawValues.Count > 0)
                {
                    var retainedCount = Math.Min(sharedProfile.RawValues.Count, result.RawValues.Count);
                    for (var index = 0; index < retainedCount; index++)
                    {
                        result.RawValues[index] = sharedProfile.RawValues[index];
                    }

                    EnsureValuesMatchIndependentInterval(rule, result.RawValues);
                    result = BuildGenerationResult(result.RawValues, decimalPlaces, result.Direction);
                }

                var reusableValues = result.RawValues
                    .Select((value, index) => Math.Round(value, decimalPlaces[index]))
                    .ToList();
                reusableValues = EnsureVisibleMeasurementVariation(
                    rule,
                    standardValue,
                    reusableValues,
                    decimalPlaces,
                    errorDecimalPlaces);
                var reusableResult = BuildGenerationResult(reusableValues, decimalPlaces, result.Direction);
                var trendError = CalculateRoundedFormulaError(
                    rule,
                    standardValue,
                    reusableValues,
                    errorDecimalPlaces);
                if (Math.Abs(trendError) <= 1e-12)
                {
                    technicalRequirementRejectCount++;
                    continue;
                }
                try
                {
                    ValidateGeneratedError(rule, standardValue, trendError);
                }
                catch (InvalidOperationException)
                {
                    technicalRequirementRejectCount++;
                    continue;
                }

                var representativeError = CalculateRepresentativeError(rule, standardValue, reusableValues);
                try
                {
                    ValidateConfiguredErrorUsage(rule, standardValue, trendError);
                }
                catch (InvalidOperationException)
                {
                    configuredIntervalRejectCount++;
                    var rejectedUsage = ResolveConfiguredErrorUsage(rule, standardValue, trendError);
                    if (rejectedUsage.HasValue)
                    {
                        minimumRejectedUsage = !minimumRejectedUsage.HasValue
                            ? rejectedUsage.Value
                            : Math.Min(minimumRejectedUsage.Value, rejectedUsage.Value);
                        maximumRejectedUsage = !maximumRejectedUsage.HasValue
                            ? rejectedUsage.Value
                            : Math.Max(maximumRejectedUsage.Value, rejectedUsage.Value);
                    }
                    minimumRejectedRepresentativeError = !minimumRejectedRepresentativeError.HasValue
                        ? representativeError
                        : Math.Min(minimumRejectedRepresentativeError.Value, representativeError);
                    maximumRejectedRepresentativeError = !maximumRejectedRepresentativeError.HasValue
                        ? representativeError
                        : Math.Max(maximumRejectedRepresentativeError.Value, representativeError);
                    continue;
                }

                if (HasDuplicateTrendError(standardValue, trendError, trendErrors))
                {
                    if (!TryAdjustForDistinctTrendError(
                        rule,
                        standardValue,
                        reusableValues,
                        decimalPlaces,
                        errorDecimalPlaces,
                        trendErrors,
                        out var adjustedValues,
                        out var adjustedTrendError))
                    {
                        duplicateRejectCount++;
                        continue;
                    }

                    reusableValues = adjustedValues;
                    reusableResult = BuildGenerationResult(reusableValues, decimalPlaces, result.Direction);
                    trendError = adjustedTrendError;
                    representativeError = CalculateRepresentativeError(rule, standardValue, reusableValues);
                }

                if (!CanAddTrendError(rule, standardValue, trendError, trendErrors))
                {
                    trendRejectCount++;
                    if (TryAdjustForDistinctTrendError(
                        rule,
                        standardValue,
                        reusableValues,
                        decimalPlaces,
                        errorDecimalPlaces,
                        trendErrors,
                        out var adjustedValues,
                        out var adjustedTrendError))
                    {
                        reusableValues = adjustedValues;
                        reusableResult = BuildGenerationResult(reusableValues, decimalPlaces, result.Direction);
                        trendError = adjustedTrendError;
                        representativeError = CalculateRepresentativeError(rule, standardValue, reusableValues);

                        AddTrendError(rule, standardValue, trendError, trendErrors);
                        session.AddTrendError(trendKey, standardValue, trendError);
                        session.Store(sharedKey, crossItemKey, reusableResult, representativeError);
                        return new GeneratedPoint(reusableResult, representativeError);
                    }

                    continue;
                }

                AddTrendError(rule, standardValue, trendError, trendErrors);
                session.AddTrendError(trendKey, standardValue, trendError);
                session.Store(
                    sharedKey,
                    crossItemKey,
                    reusableResult,
                    representativeError);
                return new GeneratedPoint(
                    reusableResult,
                    representativeError);
            }

            throw new InvalidOperationException(
                $"“{GenerationRuleValidator.ResolveRuleName(rule)}”在当前技术要求、趋势波动和小数分辨力下无法生成有效误差值。" +
                $"标准值={standardValue.ToString("G17", CultureInfo.InvariantCulture)}；" +
                $"技术要求拒绝={technicalRequirementRejectCount}，" +
                $"生成区间拒绝={configuredIntervalRejectCount}，" +
                $"趋势拒绝={trendRejectCount}，重复值拒绝={duplicateRejectCount}；" +
                "不同标准值必须产生不同误差，请扩大允许误差占用比例或提高测量值小数分辨力；" +
                $"被拒绝占用比例={FormatDiagnosticRange(minimumRejectedUsage, maximumRejectedUsage)}，" +
                $"代表误差={FormatDiagnosticRange(minimumRejectedRepresentativeError, maximumRejectedRepresentativeError)}；" +
                $"MPE={rule.FixedMpe?.ToString("G17", CultureInfo.InvariantCulture) ?? "无"}，" +
                $"误差类型={GenerationRuleValidator.ResolveGenerationErrorType(rule)}，" +
                $"量程={rule.FixedReferenceRange?.ToString("G17", CultureInfo.InvariantCulture) ?? "无"}。");
        }

        private List<double> EnsureVisibleMeasurementVariation(
            MeasurementRule rule,
            double standardValue,
            IReadOnlyList<double> values,
            IReadOnlyList<int> decimalPlaces,
            int errorDecimalPlaces)
        {
            if (values == null || values.Count <= 1 || decimalPlaces == null || decimalPlaces.Count != values.Count)
            {
                return values?.ToList() ?? new List<double>();
            }

            if (values
                .Select((value, index) => Math.Round(value, decimalPlaces[index]))
                .Distinct()
                .Count() > 1)
            {
                return values.ToList();
            }

            var resolution = Math.Pow(10, -Math.Max(0, Math.Min(15, decimalPlaces.Min())));
            for (var index = values.Count - 1; index >= 0; index--)
            {
                foreach (var direction in new[] { 1d, -1d })
                {
                    var candidate = values.ToList();
                    candidate[index] = Math.Round(values[index] + direction * resolution, decimalPlaces[index]);
                    if (Math.Abs(candidate[index] - values[index]) <= 1e-12 ||
                        candidate.Select((value, itemIndex) => Math.Round(value, decimalPlaces[itemIndex])).Distinct().Count() <= 1 ||
                        !AreMeasurementValuesValid(rule, standardValue, candidate))
                    {
                        continue;
                    }

                    var error = CalculateRoundedFormulaError(rule, standardValue, candidate, errorDecimalPlaces);
                    if (Math.Abs(error) <= 1e-12 ||
                        (rule.PositiveDirectionOnly && error < 0) ||
                        (rule.NegativeDirectionOnly && error > 0))
                    {
                        continue;
                    }

                    try
                    {
                        ValidateGeneratedError(rule, standardValue, error);
                        ValidateConfiguredErrorUsage(rule, standardValue, error);
                        return candidate;
                    }
                    catch (InvalidOperationException)
                    {
                        // Try the other representable step or another measurement cell.
                    }
                }
            }

            throw new InvalidOperationException(
                $"“{GenerationRuleValidator.ResolveRuleName(rule)}”在当前测量分辨力下无法生成不同的示值。");
        }

        private static bool AreMeasurementValuesValid(
            MeasurementRule rule,
            double standardValue,
            IReadOnlyList<double> values)
        {
            if (rule == null || values == null)
            {
                return false;
            }

            var bounds = ResolveErrorBounds(rule, standardValue);
            return values.All(value =>
            {
                var error = value - standardValue;
                if (error < bounds.lower - 1e-12 || error > bounds.upper + 1e-12)
                {
                    return false;
                }

                if (rule.MeasurementLowerBound.HasValue && value < rule.MeasurementLowerBound.Value - 1e-12)
                {
                    return false;
                }

                if (rule.MeasurementUpperBound.HasValue && value > rule.MeasurementUpperBound.Value + 1e-12)
                {
                    return false;
                }

                if (rule.PositiveDirectionOnly && error <= 1e-12)
                {
                    return false;
                }

                if (rule.NegativeDirectionOnly && error >= -1e-12)
                {
                    return false;
                }

                return true;
            });
        }

        private bool TryAdjustForDistinctTrendError(
            MeasurementRule rule,
            double standardValue,
            IReadOnlyList<double> values,
            IReadOnlyList<int> decimalPlaces,
            int errorDecimalPlaces,
            IEnumerable<TrendErrorSample> existing,
            out List<double> adjustedValues,
            out double adjustedTrendError)
        {
            adjustedValues = null;
            adjustedTrendError = 0;
            if (values == null || values.Count == 0 || decimalPlaces == null || decimalPlaces.Count != values.Count)
            {
                return false;
            }

            var minimumDecimalPlaces = decimalPlaces.Min();
            if (minimumDecimalPlaces < 0 || minimumDecimalPlaces > 15)
            {
                return false;
            }

            var resolution = Math.Pow(10, -minimumDecimalPlaces);
            var originalError = CalculateRoundedFormulaError(
                rule,
                standardValue,
                values,
                errorDecimalPlaces);
            const int maximumAdjustmentSteps = 64;
            for (var step = 1; step <= maximumAdjustmentSteps; step++)
            {
                foreach (var direction in new[] { 1d, -1d })
                {
                    var adjustment = direction * resolution * step;
                    var candidateValues = values
                        .Select((value, index) => Math.Round(value + adjustment, decimalPlaces[index]))
                        .ToList();
                    if (candidateValues
                        .Select((value, index) => Math.Abs(value - values[index]) <= 1e-12)
                        .All(isUnchanged => isUnchanged))
                    {
                        continue;
                    }

                    try
                    {
                        EnsureValuesMatchIndependentInterval(rule, candidateValues);
                        var candidateError = CalculateRoundedFormulaError(
                            rule,
                            standardValue,
                            candidateValues,
                            errorDecimalPlaces);
                        if (Math.Abs(candidateError) <= 1e-12 ||
                            HasDuplicateTrendError(standardValue, candidateError, existing) ||
                            !CanAddTrendError(rule, standardValue, candidateError, existing))
                        {
                            continue;
                        }

                        if ((rule?.PositiveDirectionOnly == true && candidateError < 0) ||
                            (rule?.NegativeDirectionOnly == true && candidateError > 0) ||
                            (_generationConfiguration?.UseSameDeviationDirection == true &&
                             Math.Sign(originalError) != Math.Sign(candidateError)))
                        {
                            continue;
                        }

                        ValidateGeneratedError(rule, standardValue, candidateError);
                        ValidateConfiguredErrorUsage(rule, standardValue, candidateError);
                        adjustedValues = candidateValues;
                        adjustedTrendError = candidateError;
                        return true;
                    }
                    catch (InvalidOperationException)
                    {
                        // Continue searching at the next representable measurement step.
                    }
                }
            }

            return false;
        }

        private static List<TrendErrorSample> LoadTrendErrors(
            MeasurementGenerationSession session,
            string trendKey,
            MeasurementRule rule)
        {
            return (session?.GetTrendErrors(trendKey) ?? new List<GeneratedTrendError>())
                .Select(item => new TrendErrorSample(
                    item.StandardValue,
                    item.Value,
                    ResolveTrendErrorUsage(rule, item.StandardValue, item.Value)))
                .ToList();
        }

        private static string FormatDiagnosticRange(double? minimum, double? maximum)
        {
            if (!minimum.HasValue || !maximum.HasValue)
            {
                return "无";
            }

            return minimum.Value.ToString("G6", CultureInfo.InvariantCulture) +
                "~" +
                maximum.Value.ToString("G6", CultureInfo.InvariantCulture);
        }

        private MeasurementRule ResolvePointRule(
            WorkbookSnapshot snapshot,
            MeasurementRule source,
            int row,
            double standardValue,
            IReadOnlyList<CellAddress> measurementCells)
        {
            var rule = MeasurementRuleCloner.Clone(source);
            var sourceFormulaScale = rule.ErrorFormula?.Scale;
            rule.FixedStandardValue = standardValue;
            rule.ManualStandardValues = new List<ManualStandardValue>();
            rule.TargetRange = new CellRange
            {
                SheetName = source.TargetRange.SheetName,
                StartRow = row,
                EndRow = row,
                StartColumn = measurementCells.Min(cell => cell.Column),
                EndColumn = measurementCells.Max(cell => cell.Column)
            };
            rule.WritableCells = CloneCellAddresses(measurementCells);

            var mapping = source.RowMappings?.FirstOrDefault(item => item?.Row == row);
            ApplyPointRange(rule.StandardValueSource, mapping?.StandardValueRange ?? SelectRangeForRow(source.StandardValueSource?.Range, row));
            ApplyPointRange(rule.AverageSource, mapping?.AverageRange ?? SelectRangeForRow(source.AverageSource?.Range, row));
            ApplyPointRange(rule.ErrorSource, mapping?.ErrorRange ?? SelectRangeForRow(source.ErrorSource?.Range, row));
            ApplyPointRange(rule.MpeSource, mapping?.TechnicalRequirementRange ?? SelectRangeForRow(source.MpeSource?.Range, row));
            ApplyPointRange(rule.RangeSource, mapping?.RangeValueRange ?? SelectRangeForRow(source.RangeSource?.Range, row));
            ApplyPointRange(rule.UncertaintySource, mapping?.UncertaintyRange ?? SelectRangeForRow(source.UncertaintySource?.Range, row));
            ApplyPointRange(rule.ResultSource, mapping?.ResultRange ?? SelectRangeForRow(source.ResultSource?.Range, row));

            if (_parameterResolver != null && snapshot != null)
            {
                rule = _parameterResolver.Apply(snapshot, new[] { rule }).Single();
            }

            rule.ErrorFormula = null;
            _structureAnalyzer.Apply(snapshot, new[] { rule });
            if (sourceFormulaScale == ErrorFormulaScale.RelativeToReferenceRange &&
                rule.ErrorFormula?.Scale == ErrorFormulaScale.RelativeToStandardValue)
            {
                rule.ErrorFormula.Scale = ErrorFormulaScale.RelativeToReferenceRange;
                rule.ErrorFormula.FormulaDividesByReferenceRange = true;
            }
            if (_parameterResolver != null && snapshot != null)
            {
                rule = _parameterResolver.Apply(snapshot, new[] { rule }).Single();
            }
            NormalizeMpeScaleForFormula(rule);
            return rule;
        }

        private static void NormalizeMpeScaleForFormula(MeasurementRule rule)
        {
            var formulaScale = rule?.ErrorFormula?.Scale;
            if (rule?.ErrorFormula?.FormulaMultipliesBy100 != true ||
                (formulaScale != ErrorFormulaScale.RelativeToStandardValue &&
                 formulaScale != ErrorFormulaScale.RelativeToReferenceRange))
            {
                return;
            }

            var pattern = MpeValuePatternCodec.Parse(rule.MpeSource?.ValuePattern);
            if (pattern == null || Math.Abs(pattern.ScaleFactor - 1d) > 1e-12)
            {
                return;
            }

            rule.FixedMpe = rule.FixedMpe / 100d;
            rule.FixedNegativeTolerance = rule.FixedNegativeTolerance / 100d;
            rule.FixedPositiveTolerance = rule.FixedPositiveTolerance / 100d;
            rule.ErrorType = formulaScale == ErrorFormulaScale.RelativeToReferenceRange
                ? ErrorType.Referenced
                : ErrorType.Relative;
            rule.MpeSource.ValuePattern = MpeValuePatternCodec.Build(
                rule.ErrorType,
                0.01d,
                rule.RequirementOperator);
        }

        private static void ApplyPointRange(ParameterSource source, CellRange range)
        {
            if (source != null && range != null)
            {
                source.Range = range;
            }
        }

        private static CellRange SelectRangeForRow(CellRange range, int row)
        {
            if (range == null || row < range.StartRow || row > range.EndRow)
            {
                return range;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = row,
                EndRow = row,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }

        private static IReadOnlyList<string> BuildGenerationWarnings(MeasurementRule rule)
        {
            if (rule?.ErrorFormula?.HasFormula == true)
            {
                return new List<string>();
            }

            return new List<string>
            {
                $"“{GenerationRuleValidator.ResolveRuleName(rule)}”"
            };
        }

    }
}
