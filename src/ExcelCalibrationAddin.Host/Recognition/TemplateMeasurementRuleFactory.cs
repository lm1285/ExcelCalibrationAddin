using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Services;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Recognition
{
    public sealed class TemplateMeasurementRuleFactory
    {
        private readonly NumberFormatInterpreter _numberFormatInterpreter;
        private readonly TemplateFieldDefinitionBuilder _templateDefinitionBuilder;

        public TemplateMeasurementRuleFactory(NumberFormatInterpreter numberFormatInterpreter)
        {
            _numberFormatInterpreter = numberFormatInterpreter;
            _templateDefinitionBuilder = new TemplateFieldDefinitionBuilder(numberFormatInterpreter);
        }

        public IReadOnlyList<MeasurementRule> BuildDraftRules(
            RecognitionResult recognition,
            IReadOnlyList<TemplateRegionMapping> mappings)
        {
            return (mappings ?? new List<TemplateRegionMapping>())
                .Select(mapping => BuildRuleFromMapping(recognition, mapping))
                .Where(rule => rule != null)
                .ToList();
        }

        public MeasurementRule BuildRuleFromMapping(RecognitionResult recognition, TemplateRegionMapping mapping)
        {
            if (mapping?.MeasurementValueRange == null)
            {
                return null;
            }

            var sheet = recognition.Snapshot.Sheets.FirstOrDefault(item => item.Name == mapping.MeasurementValueRange.SheetName);
            var formatSourceCell = sheet?.Cells.FirstOrDefault(item =>
                item.Row >= mapping.MeasurementValueRange.StartRow &&
                item.Row <= mapping.MeasurementValueRange.EndRow &&
                item.Column >= mapping.MeasurementValueRange.StartColumn &&
                item.Column <= mapping.MeasurementValueRange.EndColumn &&
                !string.IsNullOrWhiteSpace(item.NumberFormat));
            var formatRule = _numberFormatInterpreter.Interpret(formatSourceCell?.NumberFormat ?? string.Empty);

            var writableCells = ResolveWritableTemplateCells(sheet, mapping.MeasurementValueRange);

            var rule = new MeasurementRule
            {
                FieldName = mapping.ProjectName,
                FieldAlias = mapping.ProjectName,
                TargetRange = CloneRange(mapping.MeasurementValueRange),
                ErrorType = ErrorType.Absolute,
                FillMode = FillMode.Block,
                DistributionMode = DistributionMode.TruncatedNormal,
                FormatRule = formatRule,
                SetpointSource = BuildParameterSource("\u8BBE\u5B9A\u503C", mapping.SetpointValueRange),
                StandardValueSource = BuildParameterSource("\u6807\u51C6\u503C", mapping.StandardValueRange),
                AverageSource = BuildParameterSource("\u5E73\u5747\u503C", mapping.AverageValueRange),
                ErrorSource = BuildParameterSource("\u8BEF\u5DEE", mapping.ErrorValueRange),
                MpeSource = BuildParameterSource("\u6280\u672F\u8981\u6C42", mapping.TechnicalRequirementRange),
                RangeSource = BuildParameterSource("\u91CF\u7A0B", mapping.RangeValueRange),
                UncertaintySource = BuildParameterSource("\u4E0D\u786E\u5B9A\u5EA6", mapping.UncertaintyRange),
                ResultSource = BuildParameterSource("\u7ED3\u8BBA", mapping.ResultRange),
                WritableCells = writableCells,
                GroupSize = writableCells.Count > 0
                    ? writableCells.Count
                    : CountRangeCells(mapping.MeasurementValueRange)
            };
            rule.TemplateDefinition = _templateDefinitionBuilder.Build(sheet, mapping);
            return rule;
        }

        private static ParameterSource BuildParameterSource(string name, CellRange range)
        {
            if (range == null)
            {
                return null;
            }

            return new ParameterSource
            {
                Name = name,
                Range = CloneRange(range)
            };
        }

        private static CellRange CloneRange(CellRange range)
        {
            if (range == null)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                EndRow = range.EndRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }

        private static int CountRangeCells(CellRange range)
        {
            if (range == null ||
                range.EndRow < range.StartRow ||
                range.EndColumn < range.StartColumn)
            {
                return 0;
            }

            return (range.EndRow - range.StartRow + 1) * (range.EndColumn - range.StartColumn + 1);
        }

        private static List<CellAddress> ResolveWritableTemplateCells(SheetSnapshot sheet, CellRange range)
        {
            var result = new List<CellAddress>();
            if (sheet == null || range == null)
            {
                return result;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var cell in sheet.Cells
                         .Where(cell =>
                             cell.Row >= range.StartRow &&
                             cell.Row <= range.EndRow &&
                             cell.Column >= range.StartColumn &&
                             cell.Column <= range.EndColumn)
                         .OrderBy(cell => cell.Row)
                         .ThenBy(cell => cell.Column))
            {
                var effectiveRange = MergedCellLogicalRangeResolver.ResolveEffectiveRange(cell);
                if (effectiveRange.EndRow < range.StartRow ||
                    effectiveRange.StartRow > range.EndRow ||
                    effectiveRange.EndColumn < range.StartColumn ||
                    effectiveRange.StartColumn > range.EndColumn)
                {
                    continue;
                }

                var anchor = MergedCellLogicalRangeResolver.ResolveAnchorCell(sheet, cell);
                if (anchor == null || !string.IsNullOrWhiteSpace(anchor.Formula))
                {
                    continue;
                }

                var key = $"{effectiveRange.StartRow}:{effectiveRange.StartColumn}:{effectiveRange.EndRow}:{effectiveRange.EndColumn}";
                if (!seen.Add(key))
                {
                    continue;
                }

                result.Add(new CellAddress
                {
                    Row = effectiveRange.StartRow,
                    Column = effectiveRange.StartColumn
                });
            }

            return result;
        }
    }
}
