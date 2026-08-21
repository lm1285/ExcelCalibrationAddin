using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Services;

namespace ExcelCalibrationAddin.Vsto.TaskPane
{
    internal static class TaskPaneModelCloner
    {
        public static List<TemplateRegionMapping> CloneMappings(IReadOnlyList<TemplateRegionMapping> mappings)
        {
            return (mappings ?? new List<TemplateRegionMapping>())
                .Select(item => new TemplateRegionMapping
                {
                    ProjectName = item.ProjectName,
                    SectionRange = CloneRange(item.SectionRange),
                    SetpointValueRange = CloneRange(item.SetpointValueRange),
                    StandardValueRange = CloneRange(item.StandardValueRange),
                    MeasurementValueRange = CloneRange(item.MeasurementValueRange),
                    AverageValueRange = CloneRange(item.AverageValueRange),
                    ErrorValueRange = CloneRange(item.ErrorValueRange),
                    TechnicalRequirementRange = CloneRange(item.TechnicalRequirementRange),
                    UncertaintyRange = CloneRange(item.UncertaintyRange),
                    RangeValueRange = CloneRange(item.RangeValueRange),
                    ResultRange = CloneRange(item.ResultRange),
                    Notes = item.Notes
                })
                .ToList();
        }

        public static List<MeasurementRule> CloneRules(IReadOnlyList<MeasurementRule> rules)
        {
            return (rules ?? new List<MeasurementRule>())
                .Where(rule => rule != null)
                .Select(MeasurementRuleCloner.Clone)
                .ToList();
        }

        public static MeasurementRule CloneRule(MeasurementRule rule)
        {
            return MeasurementRuleCloner.Clone(rule);
        }

        public static ParameterSource BuildParameterSource(ParameterSource existing, string name, CellRange range)
        {
            if (range == null)
            {
                return CloneParameterSource(existing);
            }

            return new ParameterSource
            {
                Name = string.IsNullOrWhiteSpace(existing?.Name) ? name : existing.Name,
                Range = CloneRange(range),
                ValuePattern = existing?.ValuePattern ?? string.Empty
            };
        }

        public static ParameterSource CloneParameterSource(ParameterSource source)
        {
            if (source == null)
            {
                return null;
            }

            return new ParameterSource
            {
                Name = source.Name,
                Range = CloneRange(source.Range),
                ValuePattern = source.ValuePattern
            };
        }

        public static CellRange CloneRange(CellRange range)
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

        public static TemplateFingerprint CloneFingerprint(TemplateFingerprint fingerprint)
        {
            if (fingerprint == null)
            {
                return null;
            }

            return new TemplateFingerprint
            {
                ExactFingerprint = fingerprint.ExactFingerprint,
                FuzzyFingerprint = fingerprint.FuzzyFingerprint,
                StructureSignature = fingerprint.StructureSignature,
                Summary = fingerprint.Summary,
                Title = fingerprint.Title,
                SheetNames = new List<string>(fingerprint.SheetNames ?? new List<string>()),
                HeaderTexts = new List<string>(fingerprint.HeaderTexts ?? new List<string>())
            };
        }

        public static List<CellAddress> CloneCellAddresses(IEnumerable<CellAddress> cells)
        {
            return (cells ?? Enumerable.Empty<CellAddress>())
                .Where(cell => cell != null && cell.Row > 0 && cell.Column > 0)
                .Select(cell => new CellAddress { Row = cell.Row, Column = cell.Column })
                .ToList();
        }
    }
}
