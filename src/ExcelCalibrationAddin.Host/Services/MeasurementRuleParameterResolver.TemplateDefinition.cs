using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleParameterResolver
    {
        private static void RestoreParameterSourcesFromTemplateDefinition(MeasurementRule rule)
        {
            var regions = rule?.TemplateDefinition?.Regions ?? new List<TemplateRegionDefinition>();
            if (rule == null || regions.Count == 0) return;

            rule.SetpointSource = RestoreSource(
                rule.SetpointSource,
                "设定值",
                FindRange(regions, TemplateRegionRole.SetpointValue));
            rule.StandardValueSource = RestoreSource(
                rule.StandardValueSource,
                "标准值",
                FindRange(regions, TemplateRegionRole.StandardValue));
            rule.AverageSource = RestoreSource(
                rule.AverageSource,
                "平均值",
                FindRange(regions, TemplateRegionRole.AverageValue));
            rule.ErrorSource = RestoreSource(
                rule.ErrorSource,
                "误差",
                FindRange(regions, TemplateRegionRole.ErrorValue));
            rule.MpeSource = RestoreSource(
                rule.MpeSource,
                "技术要求",
                FindRange(regions, TemplateRegionRole.TechnicalRequirement));
            rule.RangeSource = RestoreSource(
                rule.RangeSource,
                "量程",
                FindRange(regions, TemplateRegionRole.RangeValue));
            rule.UncertaintySource = RestoreSource(
                rule.UncertaintySource,
                "不确定度",
                FindRange(regions, TemplateRegionRole.Uncertainty));
            rule.ResultSource = RestoreSource(
                rule.ResultSource,
                "结果",
                FindRange(regions, TemplateRegionRole.Result));
        }

        private static ParameterSource RestoreSource(ParameterSource source, string name, CellRange range)
        {
            if (range == null || source?.Range != null) return source;
            return new ParameterSource
            {
                Name = string.IsNullOrWhiteSpace(source?.Name) ? name : source.Name,
                Range = CloneTemplateRange(range),
                ValuePattern = source?.ValuePattern ?? string.Empty
            };
        }

        private static CellRange FindRange(
            IEnumerable<TemplateRegionDefinition> regions,
            TemplateRegionRole role)
        {
            return regions.FirstOrDefault(region => region?.Role == role)?.Range;
        }

        private static CellRange CloneTemplateRange(CellRange range)
        {
            return range == null
                ? null
                : new CellRange
                {
                    SheetName = range.SheetName,
                    StartRow = range.StartRow,
                    EndRow = range.EndRow,
                    StartColumn = range.StartColumn,
                    EndColumn = range.EndColumn
                };
        }
    }
}
