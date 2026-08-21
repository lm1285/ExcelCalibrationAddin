using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Host.ViewModels;
using ExcelCalibrationAddin.Host.Vsto;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private TaskPaneState TryGenerateFromCachedState(
            Excel.Workbook workbook,
            string workbookKey,
            MeasurementGenerationOverride generationOverride)
        {
            if (!CanUseCachedGenerationState(workbookKey))
            {
                return null;
            }

            var cachedState = _lastGenerationState;
            var rulesToWrite = VstoAddinFacade.ApplyGenerationOverride(cachedState.DraftRules, generationOverride);
            if (rulesToWrite == null || rulesToWrite.Count == 0)
            {
                return null;
            }

            var appliedConfiguration = cachedState.AppliedGenerationConfiguration ?? LoadGenerationConfiguration();
            Trace.WriteLine($"[VSTO] GenerateRandom using cached rules. Workbook={workbook?.Name}, Rules={rulesToWrite.Count}");
            _isWritingGeneratedValues = true;
            var writeWatch = Stopwatch.StartNew();
            try
            {
                // Saved task-pane rules can contain references whose values must be read from
                // the current workbook (for example, a full-scale reference range). Do not
                // bypass parameter resolution merely because the rules came from the cache.
                var writeResult = _facade.WriteRules(workbook, rulesToWrite, appliedConfiguration);
                cachedState.GenerationWarningMessages = writeResult?.WarningMessages ?? new List<string>();
            }
            finally
            {
                writeWatch.Stop();
                WriteGenerationPerformanceLog($"cachedWrite elapsed={writeWatch.ElapsedMilliseconds}ms rules={rulesToWrite.Count}");
                Trace.WriteLine($"[VSTO] GenerateRandom cached write elapsed={writeWatch.ElapsedMilliseconds}ms");
                _isWritingGeneratedValues = false;
            }

            return new TaskPaneState
            {
                WorkbookName = cachedState.WorkbookName,
                ExactFingerprint = cachedState.ExactFingerprint,
                RemoteTemplateId = cachedState.RemoteTemplateId,
                RemoteTemplateName = cachedState.RemoteTemplateName,
                RemoteStatus = cachedState.RemoteStatus,
                RemoteDetail = cachedState.RemoteDetail,
                MatchStatus = cachedState.MatchStatus,
                MatchStatusDetail = cachedState.MatchStatusDetail,
                RecognitionStatusDetail = cachedState.RecognitionStatusDetail,
                FingerprintStatusDetail = cachedState.FingerprintStatusDetail,
                LocalMatchStatusDetail = cachedState.LocalMatchStatusDetail,
                RemoteMatchStatusDetail = cachedState.RemoteMatchStatusDetail,
                DraftRuleStatusDetail = cachedState.DraftRuleStatusDetail,
                LocalTemplateStatus = cachedState.LocalTemplateStatus,
                LocalTemplateStatusDetail = cachedState.LocalTemplateStatusDetail,
                AppliedGenerationConfiguration = appliedConfiguration,
                UsesTemplateGenerationConfiguration = cachedState.UsesTemplateGenerationConfiguration,
                GenerationWarningMessages = cachedState.GenerationWarningMessages,
                Fingerprint = cachedState.Fingerprint,
                IsFeatureBlocked = cachedState.IsFeatureBlocked,
                CanGenerate = cachedState.CanGenerate,
                RecognizedFields = cachedState.RecognizedFields,
                MappingItems = cachedState.MappingItems,
                DraftRules = rulesToWrite
            };
        }

        private bool CanUseCachedGenerationState(string workbookKey)
        {
            return _lastGenerationState != null &&
                _lastGenerationState.CanGenerate &&
                !_lastGenerationState.IsFeatureBlocked &&
                _lastGenerationState.DraftRules != null &&
                _lastGenerationState.DraftRules.Count > 0 &&
                string.Equals(_lastMatchedWorkbookKey, workbookKey, StringComparison.OrdinalIgnoreCase);
        }

        private bool HasCachedWorkbookMatchState(string workbookKey)
        {
            return _lastGenerationState != null &&
                string.Equals(_lastMatchedWorkbookKey, workbookKey, StringComparison.OrdinalIgnoreCase);
        }

        private ExcelGenerationPerformanceScope EnterFastGenerationMode()
        {
            return new ExcelGenerationPerformanceScope(this.Application);
        }

        private void RememberGenerationState(string workbookKey, TaskPaneState state)
        {
            if (state == null)
            {
                return;
            }

            _lastGenerationState = state;
            _lastMatchedWorkbookKey = workbookKey;
        }

        internal void UpdateCurrentGenerationRules(
            System.Collections.Generic.IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration)
        {
            var workbook = this.Application?.ActiveWorkbook as Excel.Workbook;
            var workbookKey = ResolveWorkbookKey(workbook);
            if (_lastGenerationState == null)
            {
                _lastGenerationState = new TaskPaneState
                {
                    WorkbookName = SafeWorkbookName(workbook)
                };
            }

            _lastGenerationState.DraftRules = rules ?? new List<MeasurementRule>();
            _lastGenerationState.AppliedGenerationConfiguration = generationConfiguration ?? LoadGenerationConfiguration();
            _lastGenerationState.GenerationWarningMessages = ResolveGenerationWarnings();
            _lastGenerationState.CanGenerate = _lastGenerationState.DraftRules.Count > 0;
            _lastGenerationState.IsFeatureBlocked = false;

            if (!string.IsNullOrWhiteSpace(workbookKey))
            {
                _lastMatchedWorkbookKey = workbookKey;
            }

            RefreshRandomRangeSummary(_lastGenerationState.DraftRules, _lastGenerationState.AppliedGenerationConfiguration);
        }

        private void ApplyGlobalGenerationConfigurationToCurrentState()
        {
            if (_lastGenerationState == null || _lastGenerationState.UsesTemplateGenerationConfiguration)
            {
                return;
            }

            _lastGenerationState.AppliedGenerationConfiguration = _generationConfiguration;
        }

        private IReadOnlyList<string> ResolveGenerationWarnings()
        {
            return _lastGenerationState?.GenerationWarningMessages ?? new List<string>();
        }

    }
}
