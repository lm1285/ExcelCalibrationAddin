using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Recognition;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleDraftBuilder
    {
        private readonly NumberFormatInterpreter _numberFormatInterpreter;
        private readonly TemplateMeasurementRuleFactory _ruleFactory;
        private readonly ErrorRangeDetector _errorRangeDetector;

        private static readonly string[] MeasurementSubHeaderExcludes = { "AVG", "\u5E73\u5747", "\u5747\u503C" };
        private static readonly Regex MeasurementAttemptHeaderRegex = new Regex(@"^(\u6D4B\u91CF|\u8BFB\u6570|\u8BFB\u503C|\u793A\u503C)?(\u7B2C)?\d+(\u6B21|\u56DE|\u70B9|\u7EC4|\u53F7|#)?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly string[] StandardKeywords = FieldHeaderVocabulary.StandardValueKeywords;
        private static readonly string[] SetpointKeywords = FieldHeaderVocabulary.SetpointValueKeywords;
        private static readonly string[] ReferenceMeasurementKeywords = FieldHeaderVocabulary.ReferenceMeasurementValueKeywords;
        private static readonly string[] MeasurementKeywords = FieldHeaderVocabulary.MeasurementValueKeywords;
        private static readonly string[] RepeatedMeasurementParentKeywords = { "\u91CD\u590D\u6027", "\u7A33\u5B9A\u6027" };
        private static readonly string[] AverageKeywords = FieldHeaderVocabulary.AverageValueKeywords;
        private static readonly string[] ErrorKeywords = FieldHeaderVocabulary.ErrorValueKeywords;
        private static readonly string[] TechnicalKeywords = FieldHeaderVocabulary.TechnicalRequirementKeywords;
        private static readonly string[] UncertaintyKeywords = FieldHeaderVocabulary.UncertaintyKeywords;
        private static readonly string[] RangeKeywords = FieldHeaderVocabulary.RangeValueKeywords;
        private static readonly string[] ResultKeywords = FieldHeaderVocabulary.ResultKeywords;

        public MeasurementRuleDraftBuilder(NumberFormatInterpreter numberFormatInterpreter)
        {
            _numberFormatInterpreter = numberFormatInterpreter;
            _ruleFactory = new TemplateMeasurementRuleFactory(numberFormatInterpreter);
            _errorRangeDetector = new ErrorRangeDetector(numberFormatInterpreter);
        }

        public IReadOnlyList<TemplateRegionMapping> BuildMappings(RecognitionResult recognition)
        {
            var mappings = new List<TemplateRegionMapping>();

            foreach (var field in recognition.RecognizedFields.Where(item => item.Score >= 70).Take(10))
            {
                var sheet = recognition.Snapshot.Sheets.FirstOrDefault(item => item.Name == field.Range.SheetName);
                if (sheet == null)
                {
                    continue;
                }

                var startRow = field.Range.StartRow;
                var endRow = TrimTrailingNoteRows(sheet, startRow, FindSectionEndRow(sheet, startRow, field.Range.EndRow));
                var isRepeatability = IsRepeatabilityProject(field.Alias);

                var setpointRange = FindDataRange(sheet, startRow, endRow, SetpointKeywords);
                var referenceMeasurementRange = FindDataRange(sheet, startRow, endRow, ReferenceMeasurementKeywords);
                var standardRange = FindDataRange(sheet, startRow, endRow, StandardKeywords);
                // Legacy templates often label the only reference input as
                // 设定值. Keep that behavior when no separate standard field
                // is present; otherwise 设定值 remains its own third field.
                if (standardRange == null && referenceMeasurementRange == null)
                {
                    standardRange = setpointRange;
                }
                // Preserve the existing StandardValueRange name and formula
                // chain; an explicit standard-instrument header wins when
                // the template also contains a separate setpoint column.
                if (referenceMeasurementRange != null)
                {
                    standardRange = referenceMeasurementRange;
                }
                var measurementRange = FindDataRange(sheet, startRow, endRow, MeasurementKeywords);
                var averageRange = FindAverageRange(sheet, startRow, endRow, measurementRange);
                measurementRange = RefineRangeFromSelectedMeasurementSubHeaders(sheet, startRow, endRow, measurementRange);
                measurementRange = measurementRange ?? InferMeasurementRangeFromNumericHeaders(sheet, startRow, endRow, standardRange, averageRange);
                measurementRange = ExcludeAverageColumns(measurementRange, averageRange);
                if (isRepeatability)
                {
                    measurementRange = RefineRepeatabilityMeasurementRange(sheet, startRow, endRow, measurementRange, standardRange, averageRange);
                }

                var errorRange = FindErrorRangeByProjectTitle(
                    sheet,
                    startRow,
                    endRow,
                    field.Alias,
                    standardRange,
                    measurementRange,
                    averageRange);

                if (RangesOverlap(errorRange, measurementRange) ||
                    RangesOverlap(errorRange, averageRange) ||
                    RangesOverlap(errorRange, standardRange))
                {
                    errorRange = null;
                }

                var headerBand = InferHeaderBand(startRow, standardRange, measurementRange, averageRange, errorRange);

                standardRange = standardRange ?? InferRangeFromLayout(sheet, headerBand, endRow, StandardKeywords, measurementRange, averageRange, errorRange);
                measurementRange = measurementRange ?? InferRangeFromLayout(sheet, headerBand, endRow, MeasurementKeywords, standardRange, averageRange, errorRange);
                averageRange = averageRange ?? InferRangeFromLayout(sheet, headerBand, endRow, AverageKeywords, measurementRange, errorRange, standardRange);
                measurementRange = ExcludeAverageColumns(measurementRange, averageRange);
                // A standalone legacy "设定值 + 测量值" template continues to
                // use 设定值 as StandardValueRange. The new optional position is
                // exposed only when all three semantic fields are present.
                if (measurementRange == null ||
                    (referenceMeasurementRange == null && RangesOverlap(setpointRange, standardRange)))
                {
                    setpointRange = null;
                }
                if (errorRange == null && IsAverageAsErrorProject(field.Alias))
                {
                    errorRange = averageRange;
                }

                var technicalRange = FindDataRange(sheet, startRow, endRow, TechnicalKeywords)
                    ?? InferRangeFromLayout(sheet, headerBand, endRow, TechnicalKeywords, errorRange, averageRange, measurementRange);

                if (RangesOverlap(technicalRange, measurementRange) ||
                    RangesOverlap(technicalRange, averageRange) ||
                    RangesOverlap(technicalRange, standardRange))
                {
                    technicalRange = InferRangeFromLayout(
                        sheet,
                        headerBand,
                        endRow,
                        TechnicalKeywords,
                        averageRange,
                        measurementRange,
                        standardRange);
                }
                var uncertaintyRange = FindDataRange(sheet, startRow, endRow, UncertaintyKeywords)
                    ?? InferRangeFromLayout(sheet, headerBand, endRow, UncertaintyKeywords, technicalRange, errorRange, averageRange);
                var rangeValueRange = FindInlineParameterRegion(sheet, startRow, endRow, RangeKeywords)
                    ?? FindInlineParameterValueRange(sheet, startRow, endRow, RangeKeywords)
                    ?? FindDataRange(sheet, startRow, endRow, RangeKeywords)
                    ?? InferRangeFromLayout(sheet, headerBand, endRow, RangeKeywords, technicalRange, uncertaintyRange, errorRange);
                var resultRange = FindDataRange(sheet, startRow, endRow, ResultKeywords)
                    ?? InferRangeFromLayout(sheet, headerBand, endRow, ResultKeywords, uncertaintyRange, technicalRange, errorRange);

                var normalizedMapping = TemplateRegionMappingNormalizer.Normalize(sheet, new TemplateRegionMapping
                {
                    ProjectName = field.Alias,
                    SectionRange = new CellRange
                    {
                        SheetName = field.Range.SheetName,
                        StartRow = startRow,
                        EndRow = endRow,
                        StartColumn = 1,
                        EndColumn = InferMaxColumn(sheet)
                    },
                    SetpointValueRange = setpointRange,
                    StandardValueRange = standardRange,
                    MeasurementValueRange = measurementRange,
                    AverageValueRange = averageRange,
                    ErrorValueRange = errorRange,
                    TechnicalRequirementRange = technicalRange,
                    UncertaintyRange = uncertaintyRange,
                    RangeValueRange = rangeValueRange,
                    ResultRange = resultRange,
                    Notes = BuildNotes(sheet, startRow, endRow)
                });

                mappings.Add(normalizedMapping);
            }

            return mappings;
        }

        public IReadOnlyList<MeasurementRule> BuildDraftRules(RecognitionResult recognition)
        {
            return BuildDraftRules(recognition, BuildMappings(recognition));
        }

        public IReadOnlyList<MeasurementRule> BuildDraftRules(RecognitionResult recognition, IReadOnlyList<TemplateRegionMapping> mappings)
        {
            return _ruleFactory.BuildDraftRules(recognition, mappings);
        }

        public MeasurementRule BuildRuleFromMapping(RecognitionResult recognition, TemplateRegionMapping mapping)
        {
            return _ruleFactory.BuildRuleFromMapping(recognition, mapping);
        }

    }
}
