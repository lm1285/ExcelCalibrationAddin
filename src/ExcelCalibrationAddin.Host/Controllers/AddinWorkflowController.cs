using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Host.Services;
using ExcelCalibrationAddin.Host.Recognition;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Controllers
{
    public sealed class AddinWorkflowController
    {
        private readonly PluginWorkflowOrchestrator _orchestrator;
        private readonly MeasurementRuleDraftBuilder _draftBuilder;
        private readonly MeasurementRuleParameterResolver _parameterResolver;
        private readonly MeasurementRuleStructureAnalyzer _structureAnalyzer = new MeasurementRuleStructureAnalyzer();
        private readonly RowMappingBuilder _rowMappingBuilder = new RowMappingBuilder();
        private readonly TemplateFieldDefinitionBuilder _templateDefinitionBuilder =
            new TemplateFieldDefinitionBuilder(new ExcelCalibrationAddin.Core.Services.NumberFormatInterpreter());

        public AddinWorkflowController(
            PluginWorkflowOrchestrator orchestrator,
            MeasurementRuleDraftBuilder draftBuilder,
            MeasurementRuleParameterResolver parameterResolver)
        {
            _orchestrator = orchestrator;
            _draftBuilder = draftBuilder;
            _parameterResolver = parameterResolver;
        }

        public async Task<RecognitionAndDraftResult> RecognizeAsync()
        {
            var result = await _orchestrator.RecognizeAndMatchAsync();
            return BuildRecognitionAndDraftResult(result);
        }

        public async Task<RecognitionAndDraftResult> RecognizeDraftAsync()
        {
            var result = await _orchestrator.RecognizeAndMatchAsync();
            return BuildRecognitionAndDraftResult(result, forceDraftRules: true);
        }

        public RecognitionAndDraftResult RecognizeLocal()
        {
            var result = _orchestrator.RecognizeAndMatchLocal();
            return BuildLocalTemplateGenerationResult(result);
        }

        private RecognitionAndDraftResult BuildLocalTemplateGenerationResult(RecognitionAndSyncResult result)
        {
            IReadOnlyList<MeasurementRule> matchedRules = IsStrongEnabledLocalMatch(result.Local)
                ? result.Local.Rules
                : null;
            IReadOnlyList<TemplateRegionMapping> mappings;
            IReadOnlyList<MeasurementRule> draftRules;
            if (TryUseSavedLayout(result.Recognition?.Snapshot, matchedRules, out draftRules))
            {
                mappings = BuildMappingsFromRules(draftRules);
            }
            else
            {
                mappings = _draftBuilder.BuildMappings(result.Recognition);
                var currentLayoutRules = _draftBuilder.BuildDraftRules(result.Recognition, mappings);
                draftRules = matchedRules != null && matchedRules.Count > 0
                    ? RebaseRulesToCurrentLayout(matchedRules, currentLayoutRules)
                    : new List<MeasurementRule>();
            }
            draftRules = _parameterResolver.Apply(result.Recognition.Snapshot, draftRules);
            _structureAnalyzer.Apply(result.Recognition.Snapshot, draftRules);
            draftRules = _parameterResolver.Apply(result.Recognition.Snapshot, draftRules);
            PopulateWritableCells(result.Recognition.Snapshot, draftRules);
            _rowMappingBuilder.Apply(result.Recognition.Snapshot, draftRules);

            return new RecognitionAndDraftResult
            {
                Recognition = result.Recognition,
                Remote = result.Remote,
                Local = result.Local,
                Mappings = mappings,
                DraftRules = draftRules
            };
        }

        private bool TryUseSavedLayout(
            WorkbookSnapshot snapshot,
            IReadOnlyList<MeasurementRule> savedRules,
            out IReadOnlyList<MeasurementRule> currentRules)
        {
            currentRules = null;
            if (snapshot == null || savedRules == null || savedRules.Count == 0 ||
                savedRules.Any(rule => rule?.TemplateDefinition == null || rule.TargetRange == null))
            {
                return false;
            }

            var rules = new List<MeasurementRule>();
            foreach (var savedRule in savedRules)
            {
                var sheet = snapshot.Sheets.FirstOrDefault(item => string.Equals(
                    item.Name,
                    savedRule.TargetRange.SheetName,
                    StringComparison.OrdinalIgnoreCase));
                if (sheet == null)
                {
                    return false;
                }

                var currentDefinition = _templateDefinitionBuilder.Build(sheet, savedRule);
                if (!TemplateFieldDefinitionMatcher.IsCompatible(
                    savedRule.TemplateDefinition,
                    currentDefinition))
                {
                    return false;
                }

                var rule = CloneRule(savedRule);
                rule.TemplateDefinition = currentDefinition;
                rules.Add(rule);
            }

            currentRules = rules;
            return true;
        }

        private static IReadOnlyList<TemplateRegionMapping> BuildMappingsFromRules(
            IEnumerable<MeasurementRule> rules)
        {
            return (rules ?? Enumerable.Empty<MeasurementRule>())
                .Where(rule => rule?.TargetRange != null)
                .Select(rule => new TemplateRegionMapping
                {
                    ProjectName = rule.FieldAlias ?? rule.FieldName,
                    SectionRange = CloneRange(rule.TemplateDefinition?.SectionRange),
                    SetpointValueRange = CloneRange(rule.SetpointSource?.Range),
                    StandardValueRange = CloneRange(rule.StandardValueSource?.Range),
                    MeasurementValueRange = CloneRange(rule.TargetRange),
                    AverageValueRange = CloneRange(rule.AverageSource?.Range),
                    ErrorValueRange = CloneRange(rule.ErrorSource?.Range),
                    TechnicalRequirementRange = CloneRange(rule.MpeSource?.Range),
                    RangeValueRange = CloneRange(rule.RangeSource?.Range),
                    UncertaintyRange = CloneRange(rule.UncertaintySource?.Range),
                    ResultRange = CloneRange(rule.ResultSource?.Range)
                })
                .ToList();
        }

        private RecognitionAndDraftResult BuildRecognitionAndDraftResult(RecognitionAndSyncResult result, bool forceDraftRules = false)
        {
            IReadOnlyList<MeasurementRule> matchedRules = !forceDraftRules && IsStrongEnabledLocalMatch(result.Local)
                ? result.Local.Rules
                : null;
            if (!forceDraftRules &&
                (matchedRules == null || matchedRules.Count == 0) &&
                result.Remote?.Rules != null &&
                result.Remote.Rules.Count > 0 &&
                (!result.Remote.Status.HasValue || result.Remote.Status.Value == TemplateLifecycleStatus.Enabled))
            {
                matchedRules = result.Remote.Rules;
            }

            var mappings = _draftBuilder.BuildMappings(result.Recognition);
            var currentLayoutRules = _draftBuilder.BuildDraftRules(result.Recognition, mappings);
            var draftRules = matchedRules != null && matchedRules.Count > 0
                ? RebaseRulesToCurrentLayout(matchedRules, currentLayoutRules)
                : currentLayoutRules;
            draftRules = _parameterResolver.Apply(result.Recognition.Snapshot, draftRules);
            _structureAnalyzer.Apply(result.Recognition.Snapshot, draftRules);
            draftRules = _parameterResolver.Apply(result.Recognition.Snapshot, draftRules);
            PopulateWritableCells(result.Recognition.Snapshot, draftRules);
            _rowMappingBuilder.Apply(result.Recognition.Snapshot, draftRules);

            return new RecognitionAndDraftResult
            {
                Recognition = result.Recognition,
                Remote = result.Remote,
                Local = result.Local,
                Mappings = mappings,
                DraftRules = draftRules
            };
        }

        private static bool IsStrongEnabledLocalMatch(ExcelCalibrationAddin.Core.Repositories.CachedTemplateRule local)
        {
            return local != null &&
                local.Status == TemplateLifecycleStatus.Enabled &&
                local.MatchScore >= 100 &&
                local.Rules != null &&
                local.Rules.Count > 0;
        }

        private static MeasurementRule CloneRule(MeasurementRule rule)
        {
            if (rule == null)
            {
                return null;
            }

            return new MeasurementRule
            {
                FieldName = rule.FieldName,
                FieldAlias = rule.FieldAlias,
                TargetRange = CloneRange(rule.TargetRange),
                ErrorType = rule.ErrorType,
                FillMode = rule.FillMode,
                DistributionMode = rule.DistributionMode,
                FormatRule = rule.FormatRule,
                SetpointSource = CloneParameterSource(rule.SetpointSource),
                StandardValueSource = CloneParameterSource(rule.StandardValueSource),
                AverageSource = CloneParameterSource(rule.AverageSource),
                ErrorSource = CloneParameterSource(rule.ErrorSource),
                MpeSource = CloneParameterSource(rule.MpeSource),
                RangeSource = CloneParameterSource(rule.RangeSource),
                UncertaintySource = CloneParameterSource(rule.UncertaintySource),
                ResultSource = CloneParameterSource(rule.ResultSource),
                FixedStandardValue = rule.FixedStandardValue,
                ManualStandardValues = (rule.ManualStandardValues ?? Enumerable.Empty<ManualStandardValue>())
                    .Where(item => item != null)
                    .Select(item => new ManualStandardValue { PointIndex = item.PointIndex, Value = item.Value })
                    .ToList(),
                FixedMpe = rule.FixedMpe,
                FixedNegativeTolerance = rule.FixedNegativeTolerance,
                FixedPositiveTolerance = rule.FixedPositiveTolerance,
                RequirementOperator = rule.RequirementOperator,
                FixedReferenceRange = rule.FixedReferenceRange,
                MeasurementLowerBound = rule.MeasurementLowerBound,
                MeasurementUpperBound = rule.MeasurementUpperBound,
                WritableCells = CloneCellAddresses(rule.WritableCells),
                RowMappings = CloneRowMappings(rule.RowMappings),
                GroupSize = rule.GroupSize,
                IsEnabled = rule.IsEnabled,
                PositiveDirectionOnly = rule.PositiveDirectionOnly,
                NegativeDirectionOnly = rule.NegativeDirectionOnly,
                GenerationCoefficientOverride = CloneCoefficientOverride(rule.GenerationCoefficientOverride),
                ErrorFormula = CloneErrorFormula(rule.ErrorFormula),
                TemplateDefinition = TemplateDefinitionCloner.Clone(rule.TemplateDefinition)
            };
        }

        private static List<MeasurementRule> RebaseRulesToCurrentLayout(
            IReadOnlyList<MeasurementRule> savedRules,
            IReadOnlyList<MeasurementRule> currentLayoutRules)
        {
            return (savedRules ?? Enumerable.Empty<MeasurementRule>())
                .Where(rule => rule != null)
                .Select(savedRule => RebaseRuleToCurrentLayout(
                    savedRule,
                    (currentLayoutRules ?? Array.Empty<MeasurementRule>())
                        .FirstOrDefault(currentRule => SameRuleName(savedRule, currentRule))))
                .ToList();
        }

        private static MeasurementRule RebaseRuleToCurrentLayout(
            MeasurementRule savedRule,
            MeasurementRule currentLayoutRule)
        {
            if (currentLayoutRule == null)
            {
                return CloneRule(savedRule);
            }

            if (!TemplateFieldDefinitionMatcher.IsCompatible(
                savedRule?.TemplateDefinition,
                currentLayoutRule.TemplateDefinition))
            {
                return UseCurrentRuleWithSavedGenerationSettings(savedRule, currentLayoutRule);
            }

            var rule = CloneRule(savedRule);

            rule.TargetRange = CloneRange(currentLayoutRule.TargetRange);
            rule.SetpointSource = BuildParameterSource(rule.SetpointSource, currentLayoutRule.SetpointSource?.Range);
            rule.StandardValueSource = BuildParameterSource(rule.StandardValueSource, currentLayoutRule.StandardValueSource?.Range);
            rule.AverageSource = BuildParameterSource(rule.AverageSource, currentLayoutRule.AverageSource?.Range);
            rule.ErrorSource = BuildParameterSource(rule.ErrorSource, currentLayoutRule.ErrorSource?.Range);
            rule.MpeSource = BuildParameterSource(rule.MpeSource, currentLayoutRule.MpeSource?.Range);
            rule.RangeSource = BuildParameterSource(rule.RangeSource, currentLayoutRule.RangeSource?.Range);
            rule.UncertaintySource = BuildParameterSource(rule.UncertaintySource, currentLayoutRule.UncertaintySource?.Range);
            rule.ResultSource = BuildParameterSource(rule.ResultSource, currentLayoutRule.ResultSource?.Range);
            rule.WritableCells = CloneCellAddresses(currentLayoutRule.WritableCells);
            rule.GroupSize = currentLayoutRule.GroupSize;
            rule.TemplateDefinition = TemplateDefinitionCloner.Clone(currentLayoutRule.TemplateDefinition);
            return rule;
        }

        private static MeasurementRule UseCurrentRuleWithSavedGenerationSettings(
            MeasurementRule savedRule,
            MeasurementRule currentLayoutRule)
        {
            var current = CloneRule(currentLayoutRule);
            if (savedRule == null)
            {
                return current;
            }

            current.IsEnabled = savedRule.IsEnabled;
            current.FillMode = savedRule.FillMode;
            current.DistributionMode = savedRule.DistributionMode;
            current.PositiveDirectionOnly = savedRule.PositiveDirectionOnly;
            current.NegativeDirectionOnly = savedRule.NegativeDirectionOnly;
            current.GenerationCoefficientOverride = CloneCoefficientOverride(savedRule.GenerationCoefficientOverride);
            if ((savedRule.ManualStandardValues ?? new List<ManualStandardValue>()).Count > 0)
            {
                current.ManualStandardValues = (savedRule.ManualStandardValues ?? new List<ManualStandardValue>())
                    .Where(item => item != null)
                    .Select(item => new ManualStandardValue { PointIndex = item.PointIndex, Value = item.Value })
                    .ToList();
                current.FixedStandardValue = current.ManualStandardValues
                    .Where(item => item.Value.HasValue)
                    .OrderBy(item => item.PointIndex)
                    .Select(item => item.Value)
                    .FirstOrDefault();
                current.MeasurementLowerBound = savedRule.MeasurementLowerBound;
                current.MeasurementUpperBound = savedRule.MeasurementUpperBound;
            }
            return current;
        }

        private static bool SameRuleName(MeasurementRule left, MeasurementRule right)
        {
            return string.Equals(NormalizeRuleName(left?.FieldAlias ?? left?.FieldName),
                NormalizeRuleName(right?.FieldAlias ?? right?.FieldName),
                StringComparison.Ordinal);
        }

        private static string NormalizeRuleName(string value)
        {
            return new string((value ?? string.Empty)
                .Where(character => !char.IsWhiteSpace(character) && character != '、' && character != ':' && character != '：')
                .ToArray())
                .ToUpperInvariant();
        }

        private static MeasurementGenerationCoefficientOverride CloneCoefficientOverride(MeasurementGenerationCoefficientOverride value)
        {
            if (value == null)
            {
                return null;
            }

            return new MeasurementGenerationCoefficientOverride
            {
                NegativeMinimumCoefficient = value.NegativeMinimumCoefficient,
                NegativeMaximumCoefficient = value.NegativeMaximumCoefficient,
                PositiveMinimumCoefficient = value.PositiveMinimumCoefficient,
                PositiveMaximumCoefficient = value.PositiveMaximumCoefficient,
                AbsoluteMinimumCoefficient = value.AbsoluteMinimumCoefficient,
                AbsoluteMaximumCoefficient = value.AbsoluteMaximumCoefficient
            };
        }

        private static ErrorFormulaInfo CloneErrorFormula(ErrorFormulaInfo info)
        {
            if (info == null)
            {
                return null;
            }

            return new ErrorFormulaInfo
            {
                HasFormula = info.HasFormula,
                Formula = info.Formula,
                ReferencesMeasurement = info.ReferencesMeasurement,
                ReferencesStandardValue = info.ReferencesStandardValue,
                ReferencesAverage = info.ReferencesAverage,
                AverageFormulaResolved = info.AverageFormulaResolved,
                AverageFormula = info.AverageFormula,
                TechnicalRequirementFormulaResolved = info.TechnicalRequirementFormulaResolved,
                TechnicalRequirementFormula = info.TechnicalRequirementFormula,
                UncertaintyFormulaResolved = info.UncertaintyFormulaResolved,
                UncertaintyFormula = info.UncertaintyFormula,
                ResultFormulaResolved = info.ResultFormulaResolved,
                ResultFormula = info.ResultFormula,
                Scale = info.Scale,
                FormulaMultipliesBy100 = info.FormulaMultipliesBy100,
                FormulaDividesByReferenceRange = info.FormulaDividesByReferenceRange
            };
        }

        private static ParameterSource BuildParameterSource(ParameterSource existing, CellRange range)
        {
            if (range == null)
            {
                return CloneParameterSource(existing);
            }

            return new ParameterSource
            {
                Name = existing?.Name ?? string.Empty,
                Range = CloneRange(range),
                ValuePattern = existing?.ValuePattern ?? string.Empty
            };
        }

        private static ParameterSource CloneParameterSource(ParameterSource source)
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

        private static int CountCells(CellRange range)
        {
            if (range == null ||
                range.EndRow < range.StartRow ||
                range.EndColumn < range.StartColumn)
            {
                return 0;
            }

            return (range.EndRow - range.StartRow + 1) * (range.EndColumn - range.StartColumn + 1);
        }

        private static void PopulateWritableCells(WorkbookSnapshot snapshot, IReadOnlyList<MeasurementRule> rules)
        {
            if (snapshot == null || rules == null)
            {
                return;
            }

            foreach (var rule in rules.Where(item => item?.TargetRange != null))
            {
                var cells = ResolveWritableCells(snapshot, rule.TargetRange);
                if (cells.Count == 0)
                {
                    continue;
                }

                rule.WritableCells = cells;
                rule.GroupSize = cells.Count;
            }
        }

        private static List<CellAddress> ResolveWritableCells(WorkbookSnapshot snapshot, CellRange range)
        {
            var result = new List<CellAddress>();
            if (snapshot == null || range == null)
            {
                return result;
            }

            var sheet = snapshot.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, range.SheetName, StringComparison.OrdinalIgnoreCase));
            var cellLookup = sheet?.Cells.ToDictionary(cell => BuildCellKey(cell.Row, cell.Column), cell => cell);

            for (var row = range.StartRow; row <= range.EndRow; row++)
            {
                for (var column = range.StartColumn; column <= range.EndColumn; column++)
                {
                    CellMeta cell = null;
                    cellLookup?.TryGetValue(BuildCellKey(row, column), out cell);
                    if (cell?.MergeRange != null &&
                        (cell.MergeRange.StartRow != row || cell.MergeRange.StartColumn != column))
                    {
                        continue;
                    }

                    result.Add(new CellAddress { Row = row, Column = column });
                }
            }

            return result;
        }

        private static string BuildCellKey(int row, int column)
        {
            return row + ":" + column;
        }

        private static List<CellAddress> CloneCellAddresses(IEnumerable<CellAddress> cells)
        {
            return (cells ?? Enumerable.Empty<CellAddress>())
                .Where(cell => cell != null && cell.Row > 0 && cell.Column > 0)
                .Select(cell => new CellAddress { Row = cell.Row, Column = cell.Column })
                .ToList();
        }

        private static List<MeasurementRowMapping> CloneRowMappings(IEnumerable<MeasurementRowMapping> mappings)
        {
            return (mappings ?? Enumerable.Empty<MeasurementRowMapping>())
                .Where(item => item != null)
                .Select(item => new MeasurementRowMapping
                {
                    Row = item.Row,
                    SetpointValueRange = CloneRange(item.SetpointValueRange),
                    StandardValueRange = CloneRange(item.StandardValueRange),
                    MeasurementCells = CloneCellAddresses(item.MeasurementCells),
                    AverageRange = CloneRange(item.AverageRange),
                    ErrorRange = CloneRange(item.ErrorRange),
                    TechnicalRequirementRange = CloneRange(item.TechnicalRequirementRange),
                    RangeValueRange = CloneRange(item.RangeValueRange),
                    UncertaintyRange = CloneRange(item.UncertaintyRange),
                    ResultRange = CloneRange(item.ResultRange),
                    IsComplete = item.IsComplete,
                    StatusMessage = item.StatusMessage
                })
                .ToList();
        }

        public GenerationWriteResult Write(IReadOnlyList<MeasurementRule> rules)
        {
            return _orchestrator.WriteGeneration(rules);
        }

        public GenerationWriteResult Write(IReadOnlyList<MeasurementRule> rules, GenerationConfiguration generationConfiguration)
        {
            return _orchestrator.WriteGeneration(rules, generationConfiguration);
        }

        public GenerationWriteResult WriteResolved(IReadOnlyList<MeasurementRule> rules, GenerationConfiguration generationConfiguration)
        {
            return _orchestrator.WriteResolvedGeneration(rules, generationConfiguration);
        }

        public GenerationWriteResult WritePreResolved(IReadOnlyList<MeasurementRule> rules, GenerationConfiguration generationConfiguration)
        {
            return _orchestrator.WritePreResolvedGeneration(rules, generationConfiguration);
        }
    }

    public sealed class RecognitionAndDraftResult
    {
        public RecognitionResult Recognition { get; set; }
        public SyncTemplateResult Remote { get; set; }
        public ExcelCalibrationAddin.Core.Repositories.CachedTemplateRule Local { get; set; }
        public IReadOnlyList<MeasurementRule> DraftRules { get; set; }
        public IReadOnlyList<TemplateRegionMapping> Mappings { get; set; }
    }
}
