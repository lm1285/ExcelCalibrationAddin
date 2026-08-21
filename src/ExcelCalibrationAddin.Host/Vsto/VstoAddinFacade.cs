using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Controllers;
using ExcelCalibrationAddin.Host.Generation;
using ExcelCalibrationAddin.Host.Interop;
using ExcelCalibrationAddin.Host.Services;
using ExcelCalibrationAddin.Host.UseCases;
using ExcelCalibrationAddin.Host.ViewModels;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Host.Vsto
{
    public sealed partial class VstoAddinFacade : IDisposable
    {
        private readonly PluginBootstrapper _bootstrapper;

        public VstoAddinFacade(PluginConfiguration configuration)
        {
            _bootstrapper = new PluginBootstrapper(configuration);
        }

        public void Dispose()
        {
            _bootstrapper.Dispose();
        }

        public async Task<TaskPaneState> RecognizeAsync(dynamic workbook, Action<int, string> progressReporter = null)
        {
            RecognitionProgress.SetReporter(progressReporter);

            try
            {
                Trace.WriteLine($"[Host] RecognizeAsync enter. Workbook={workbook?.Name}");
                AddinWorkflowController controller = CreateController(workbook);
                RecognitionAndDraftResult result = await controller.RecognizeAsync();
                var mappings = result.Mappings;

                Trace.WriteLine(
                    $"[Host] RecognizeAsync exit. " +
                    $"Fields={result.Recognition.RecognizedFields.Count}, " +
                    $"RemoteFound={result.Remote.Found}, MatchScore={result.Remote.MatchScore:F0}, " +
                    $"Mappings={mappings.Count}");

                var state = BuildTaskPaneState(
                    result,
                    result.DraftRules,
                    HasMatchedGenerationRules(result),
                    ResolveAppliedGenerationConfiguration(result),
                    HasTemplateGenerationConfiguration(result));
                _ = new DiagnosticPackageService().AppendSummaryAsync("recognition", new
                {
                    state.WorkbookName,
                    state.ExactFingerprint,
                    state.MatchStatus,
                    rule_count = state.DraftRules?.Count ?? 0,
                    mapping_count = state.MappingItems?.Count ?? 0,
                    field_regions = state.MappingItems,
                    row_mappings = state.DraftRules?.Select(rule => new
                    {
                        name = rule?.FieldAlias ?? rule?.FieldName,
                        rows = rule?.RowMappings
                    }),
                    formulas = state.DraftRules?.Select(rule => new
                    {
                        name = rule?.FieldAlias ?? rule?.FieldName,
                        rule?.ErrorFormula
                    })
                });
                return state;
            }
            finally
            {
                RecognitionProgress.ClearReporter();
            }
        }

        public async Task<TaskPaneState> RecognizeDraftAsync(dynamic workbook, Action<int, string> progressReporter = null)
        {
            RecognitionProgress.SetReporter(progressReporter);

            try
            {
                Trace.WriteLine($"[Host] RecognizeDraftAsync enter. Workbook={workbook?.Name}");
                var controller = CreateController(workbook);
                var result = await controller.RecognizeDraftAsync();
                var state = BuildTaskPaneState(
                    result,
                    result.DraftRules,
                    false,
                    ResolveCurrentGenerationConfiguration(),
                    false);
                Trace.WriteLine(
                    $"[Host] RecognizeDraftAsync exit. Workbook={state.WorkbookName}, " +
                    $"Rules={state.DraftRules?.Count ?? 0}, Mappings={state.MappingItems?.Count ?? 0}");
                return state;
            }
            finally
            {
                RecognitionProgress.ClearReporter();
            }
        }

        public Task<TaskPaneState> GenerateFromTemplateLibraryAsync(
            dynamic workbook,
            Action<int, string> progressReporter = null,
            MeasurementGenerationOverride generationOverride = null)
        {
            RecognitionProgress.SetReporter(progressReporter);

            try
            {
                Trace.WriteLine($"[Host] GenerateFromTemplateLibrary enter. Workbook={workbook?.Name}");
                var controller = CreateController(workbook);
                var recognition = controller.RecognizeLocal();
                if (!HasMatchedGenerationRules(recognition))
                {
                    return Task.FromResult(BuildTaskPaneState(recognition, recognition.DraftRules, false));
                }

                var appliedConfiguration = ResolveAppliedGenerationConfiguration(recognition);
                var rulesToWrite = ApplyGenerationOverride(recognition.DraftRules, generationOverride);
                var writeResult = controller.WriteResolved(rulesToWrite, appliedConfiguration);
                var state = BuildTaskPaneState(
                    recognition,
                    rulesToWrite,
                    true,
                    appliedConfiguration,
                    HasTemplateGenerationConfiguration(recognition));
                state.GenerationWarningMessages = writeResult?.WarningMessages ?? new List<string>();
                return Task.FromResult(state);
            }
            finally
            {
                RecognitionProgress.ClearReporter();
            }
        }

        public Task<TaskPaneState> MatchTemplateLibraryAsync(dynamic workbook, Action<int, string> progressReporter = null)
        {
            RecognitionProgress.SetReporter(progressReporter);

            try
            {
                Trace.WriteLine($"[Host] MatchTemplateLibrary enter. Workbook={workbook?.Name}");
                TaskPaneState fallbackState = null;
                foreach (var sheetName in ResolveAutomaticMatchSheetNames(workbook))
                {
                    var controller = CreateController(workbook, sheetName);
                    var recognition = controller.RecognizeLocal();
                    var canGenerate = HasMatchedGenerationRules(recognition);
                    var state = BuildTaskPaneState(
                        recognition,
                        recognition.DraftRules,
                        canGenerate,
                        ResolveAppliedGenerationConfiguration(recognition),
                        HasTemplateGenerationConfiguration(recognition));
                    fallbackState = fallbackState ?? state;
                    Trace.WriteLine(
                        $"[Host] MatchTemplateLibrary attempt. Workbook={state.WorkbookName}, " +
                        $"Sheet={sheetName}, Fingerprint={state.ExactFingerprint}, " +
                        $"CanGenerate={state.CanGenerate}, Rules={state.DraftRules?.Count ?? 0}");
                    if (canGenerate)
                    {
                        return Task.FromResult(state);
                    }
                }

                return Task.FromResult(fallbackState ?? new TaskPaneState
                {
                    WorkbookName = Convert.ToString(workbook?.Name) ?? string.Empty
                });
            }
            finally
            {
                RecognitionProgress.ClearReporter();
            }
        }

        private IReadOnlyList<string> ResolveAutomaticMatchSheetNames(dynamic workbook)
        {
            var workbookSheetNames = new List<string>();
            foreach (var worksheet in workbook.Worksheets)
            {
                var name = Convert.ToString(worksheet.Name) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    workbookSheetNames.Add(name);
                }
            }

            var templateSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var hasTemplateWithoutSheetMetadata = false;
            var enabledTemplates = _bootstrapper.LocalTemplateRuleCacheRepository.ListSavedTemplates()
                .Where(item => item != null &&
                    item.Status == TemplateLifecycleStatus.Enabled &&
                    item.Rules != null &&
                    item.Rules.Count > 0)
                .ToList();
            foreach (var template in enabledTemplates)
            {
                var names = (template.Fingerprint?.SheetNames ?? new List<string>())
                    .Concat(template.Rules
                        .Where(rule => rule?.TargetRange != null)
                        .Select(rule => rule.TargetRange.SheetName))
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .ToList();
                hasTemplateWithoutSheetMetadata = hasTemplateWithoutSheetMetadata || names.Count == 0;
                foreach (var name in names)
                {
                    templateSheetNames.Add(name);
                }
            }

            var candidates = AutomaticMatchSheetSelector.Select(
                workbookSheetNames,
                templateSheetNames,
                enabledTemplates.Count > 0,
                hasTemplateWithoutSheetMetadata);
            Trace.WriteLine(
                $"[Host] Auto match preflight. WorkbookSheets={workbookSheetNames.Count}, " +
                $"EnabledTemplates={enabledTemplates.Count}, CandidateSheets={candidates.Count}");
            return candidates;
        }

        public GenerationWriteResult WriteRules(dynamic workbook, IReadOnlyList<MeasurementRule> rules, GenerationConfiguration generationConfiguration = null)
        {
            Trace.WriteLine($"[Host] WriteRules enter. Workbook={workbook?.Name}, Rules={rules?.Count ?? 0}");
            var controller = CreateController(workbook);
            GenerationWriteResult result;
            if (generationConfiguration == null)
            {
                result = controller.Write(rules);
            }
            else
            {
                result = controller.Write(rules, generationConfiguration);
            }
            Trace.WriteLine("[Host] WriteRules exit.");
            return result;
        }

        public SampleDataCaptureResult SaveSampleData(dynamic workbook, string templateFingerprint, IReadOnlyList<MeasurementRule> rules, ISet<string> selectedNames = null)
        {
            if (workbook == null) throw new InvalidOperationException("请先打开一个工作 Excel。");
            var provider = new ExcelInteropSnapshotProvider(workbook);
            var result = _bootstrapper.CreateSampleDataUseCase(provider).CaptureAndSave(templateFingerprint, rules, selectedNames);
            Trace.WriteLine($"[Host] Sample data saved. Fingerprint={templateFingerprint}, Items={result.SavedItemCount}, Skipped={result.SkippedItems.Count}");
            return result;
        }

        public IReadOnlyList<SampleDataVersion> ListSampleDataVersions(string templateFingerprint) => _bootstrapper.LocalTemplateRuleCacheRepository.ListSampleDataVersions(templateFingerprint);
        public bool DeleteSampleDataVersion(long versionId) => _bootstrapper.LocalTemplateRuleCacheRepository.DeleteSampleDataVersion(versionId);

        public GenerationWriteResult WritePreResolvedRules(dynamic workbook, IReadOnlyList<MeasurementRule> rules, GenerationConfiguration generationConfiguration)
        {
            Trace.WriteLine($"[Host] WritePreResolvedRules enter. Workbook={workbook?.Name}, Rules={rules?.Count ?? 0}");
            var controller = CreateController(workbook);
            var result = controller.WritePreResolved(rules, generationConfiguration ?? ResolveCurrentGenerationConfiguration());
            Trace.WriteLine("[Host] WritePreResolvedRules exit.");
            return result;
        }

        public void VerifyFormulaResults(dynamic workbook, IReadOnlyList<MeasurementRule> rules)
        {
            var formulaRules = (rules ?? new List<MeasurementRule>())
                .Where(rule => rule?.ErrorSource?.Range != null &&
                               (rule.ErrorFormula?.HasFormula == true ||
                                GenerationRuleValidator.IsRepeatabilityRule(rule)))
                .ToList();
            if (formulaRules.Count == 0)
            {
                return;
            }

            var verificationRanges = formulaRules
                .Select(rule => rule.ErrorSource.Range)
                .Concat(formulaRules
                    .Where(GenerationRuleValidator.IsRepeatabilityRule)
                    .Select(rule => rule.TargetRange)
                    .Where(GenerationRuleValidator.HasValidRange))
                .ToList();
            var snapshot = new ExcelInteropSnapshotProvider(workbook).Capture(verificationRanges);
            new FormulaResultVerifier().Verify(snapshot, formulaRules);
        }

        public TemplateSaveResult SaveTemplate(
            string templateName,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration = null,
            bool createNew = false,
            TemplateDirectoryMetadata directoryMetadata = null,
            string targetRemoteTemplateId = null)
        {
            Trace.WriteLine($"[Host] SaveTemplate enter. Name={templateName}, Rules={rules?.Count ?? 0}");
            return _bootstrapper.TemplateSaveService.Save(
                templateName,
                fingerprint,
                rules,
                generationConfiguration,
                createNew,
                directoryMetadata,
                targetRemoteTemplateId);
        }

        public Task<TemplateSaveResult> SaveTemplateAsync(
            string templateName,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration = null,
            bool createNew = false,
            TemplateDirectoryMetadata directoryMetadata = null,
            string targetRemoteTemplateId = null)
        {
            Trace.WriteLine($"[Host] SaveTemplateAsync enter. Name={templateName}, Rules={rules?.Count ?? 0}");
            return _bootstrapper.TemplateSaveService.SaveAsync(
                templateName,
                fingerprint,
                rules,
                generationConfiguration,
                createNew,
                directoryMetadata,
                targetRemoteTemplateId);
        }

        public IReadOnlyList<SavedTemplateInfo> ListSavedTemplates()
        {
            return _bootstrapper.LocalTemplateRuleCacheRepository.ListSavedTemplates()
                .Select(item => new SavedTemplateInfo
                {
                    RemoteTemplateId = item.RemoteTemplateId,
                    TemplateName = item.TemplateName,
                    DirectoryMetadata = item.DirectoryMetadata,
                    ExactFingerprint = item.ExactFingerprint,
                    RuleCount = item.Rules?.Count ?? 0,
                    UpdatedAt = item.UpdatedAt,
                    Status = item.Status,
                    LocalSyncStatus = item.LocalSyncStatus,
                    SyncError = item.SyncError,
                    HasRemoteConflict = item.HasRemoteConflict,
                    HasGenerationConfigurationOverride = item.GenerationConfiguration != null
                })
                .ToList();
        }

        public TaskPaneState BuildSavedTemplateEditorState(string exactFingerprint, string remoteTemplateId = null)
        {
            var template = string.IsNullOrWhiteSpace(remoteTemplateId)
                ? _bootstrapper.LocalTemplateRuleCacheRepository.FindByExactFingerprint(exactFingerprint)
                : _bootstrapper.LocalTemplateRuleCacheRepository.FindByRemoteTemplateId(remoteTemplateId);
            if (template == null)
            {
                return null;
            }

            var rules = template.Rules ?? new List<MeasurementRule>();
            var fingerprint = template.Fingerprint ?? new TemplateFingerprint();
            fingerprint.ExactFingerprint = string.IsNullOrWhiteSpace(fingerprint.ExactFingerprint)
                ? template.ExactFingerprint
                : fingerprint.ExactFingerprint;
            return new TaskPaneState
            {
                WorkbookName = string.Empty,
                ExactFingerprint = template.ExactFingerprint,
                RemoteTemplateId = template.RemoteTemplateId,
                RemoteTemplateName = template.TemplateName,
                RemoteStatus = "Matched",
                RemoteDetail = "已从模板库加载，可直接编辑",
                MatchStatus = "Normal",
                MatchStatusDetail = "模板库编辑模式",
                LocalTemplateStatus = template.Status,
                LocalTemplateStatusDetail = $"本地模板状态: {TranslateTemplateStatus(template.Status)}",
                Fingerprint = fingerprint,
                AppliedGenerationConfiguration = template.GenerationConfiguration,
                UsesTemplateGenerationConfiguration = template.GenerationConfiguration != null,
                IsFeatureBlocked = false,
                CanGenerate = rules.Count > 0,
                MappingItems = BuildMappingsFromRules(rules),
                DraftRules = rules
            };
        }

        private static IReadOnlyList<TemplateRegionMapping> BuildMappingsFromRules(IEnumerable<MeasurementRule> rules)
        {
            return (rules ?? Enumerable.Empty<MeasurementRule>())
                .Where(rule => rule != null)
                .Select(rule => new TemplateRegionMapping
                {
                    ProjectName = string.IsNullOrWhiteSpace(rule.FieldAlias) ? rule.FieldName : rule.FieldAlias,
                    SectionRange = CloneSavedTemplateRange(rule.TemplateDefinition?.SectionRange),
                    SetpointValueRange = CloneSavedTemplateRange(rule.SetpointSource?.Range),
                    StandardValueRange = CloneSavedTemplateRange(rule.StandardValueSource?.Range),
                    MeasurementValueRange = CloneSavedTemplateRange(rule.TargetRange),
                    AverageValueRange = CloneSavedTemplateRange(rule.AverageSource?.Range),
                    ErrorValueRange = CloneSavedTemplateRange(rule.ErrorSource?.Range),
                    TechnicalRequirementRange = CloneSavedTemplateRange(rule.MpeSource?.Range),
                    RangeValueRange = CloneSavedTemplateRange(rule.RangeSource?.Range),
                    UncertaintyRange = CloneSavedTemplateRange(rule.UncertaintySource?.Range),
                    ResultRange = CloneSavedTemplateRange(rule.ResultSource?.Range)
                })
                .ToList();
        }

        private static CellRange CloneSavedTemplateRange(CellRange range)
        {
            return range == null ? null : new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                EndRow = range.EndRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }

        public bool ResolveTemplateConflict(string exactFingerprint, TemplateConflictResolutionAction action, string saveAsTemplateName = null)
        {
            return _bootstrapper.LocalTemplateRuleCacheRepository.ResolveTemplateConflict(exactFingerprint, action, saveAsTemplateName);
        }

        public bool UpdateSavedTemplateStatus(string exactFingerprint, TemplateLifecycleStatus status)
        {
            return _bootstrapper.LocalTemplateRuleCacheRepository.UpdateTemplateStatus(exactFingerprint, status);
        }

        public bool DeleteSavedTemplate(string exactFingerprint)
        {
            return _bootstrapper.LocalTemplateRuleCacheRepository.DeleteTemplate(exactFingerprint);
        }

        public int UploadPendingTemplates()
        {
            return _bootstrapper.TemplatePendingUploadService.UploadPendingTemplates();
        }

        public Task<TemplateSyncRunResult> SyncTemplatesAsync()
        {
            return _bootstrapper.CreateSyncTemplateUseCase().SyncAsync();
        }

        public DateTime? GetLastSuccessfulSyncUtc()
        {
            return _bootstrapper.CreateSyncTemplateUseCase().GetLastSuccessfulSyncUtc();
        }

        public bool ExportDiagnostics(string outputPath, TaskPaneState state)
        {
            var input = new DiagnosticPackageInput
            {
                WorkbookName = state?.WorkbookName ?? string.Empty,
                TemplateName = state?.RemoteTemplateName ?? string.Empty,
                Fingerprint = state?.ExactFingerprint ?? string.Empty,
                MatchResultJson = JsonConvert.SerializeObject(new
                {
                    state?.MatchStatus,
                    state?.MatchStatusDetail,
                    state?.RemoteStatus,
                    state?.RemoteDetail,
                    state?.LocalTemplateStatus,
                    state?.LocalTemplateStatusDetail
                }, Formatting.Indented),
                RecognitionResultJson = JsonConvert.SerializeObject(new
                {
                    fields = state?.RecognizedFields ?? new List<RecognizedField>(),
                    mappings = state?.MappingItems ?? new List<TemplateRegionMapping>(),
                    rules = (state?.DraftRules ?? new List<MeasurementRule>()).Select(rule => new
                    {
                        name = rule?.FieldAlias ?? rule?.FieldName,
                        rule?.RowMappings,
                        rule?.ErrorFormula,
                        requirement_pattern = rule?.MpeSource?.ValuePattern
                    })
                }, Formatting.Indented),
                SummaryJson = JsonConvert.SerializeObject(new
                {
                    state?.WorkbookName,
                    state?.ExactFingerprint,
                    state?.DraftRuleStatusDetail,
                    rule_count = state?.DraftRules?.Count ?? 0,
                    mapping_count = state?.MappingItems?.Count ?? 0,
                    field_regions = state?.MappingItems,
                    row_mapping_status = state?.DraftRules?.Select(rule => new
                    {
                        name = rule?.FieldAlias ?? rule?.FieldName,
                        total = rule?.RowMappings?.Count ?? 0,
                        complete = rule?.RowMappings?.Count(item => item?.IsComplete == true) ?? 0
                    }),
                    formula_status = state?.DraftRules?.Select(rule => new
                    {
                        name = rule?.FieldAlias ?? rule?.FieldName,
                        error = rule?.ErrorFormula?.HasFormula == true,
                        requirement = rule?.ErrorFormula?.TechnicalRequirementFormulaResolved == true,
                        result = rule?.ErrorFormula?.ResultFormulaResolved == true
                    }),
                    state?.LocalTemplateStatus
                }, Formatting.Indented)
            };
            new DiagnosticPackageService().Export(outputPath, input);
            return true;
        }

        public GenerationConfiguration GetSavedTemplateGenerationConfiguration(string exactFingerprint)
        {
            var template = _bootstrapper.LocalTemplateRuleCacheRepository.FindByExactFingerprint(exactFingerprint);
            return template?.GenerationConfiguration;
        }

        public bool UpdateSavedTemplateGenerationConfiguration(string exactFingerprint, GenerationConfiguration generationConfiguration)
        {
            return _bootstrapper.LocalTemplateRuleCacheRepository.UpdateTemplateGenerationConfiguration(exactFingerprint, generationConfiguration);
        }

    }
}
