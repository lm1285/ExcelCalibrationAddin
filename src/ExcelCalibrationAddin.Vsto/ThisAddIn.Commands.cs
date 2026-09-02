using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.ViewModels;
using ExcelCalibrationAddin.Host.UseCases;
using ExcelCalibrationAddin.Host.Vsto;
using ExcelCalibrationAddin.Vsto.TaskPane;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        internal async Task<bool> LoginToCloudAsync()
        {
            var configuration = new ConfigurationLoader().Load(_configPath);
            using (var dialog = new CloudLoginDialog())
            {
                if (dialog.ShowDialog(GetExcelMainWindow()) != DialogResult.OK)
                {
                    return false;
                }

                var endpoint = configuration.Backend.BaseUrl.TrimEnd('/') + "/api/auth/login";
                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) })
                using (var content = new StringContent(
                    new JavaScriptSerializer().Serialize(new { username = dialog.Username, password = dialog.Password }),
                    Encoding.UTF8,
                    "application/json"))
                using (var response = await client.PostAsync(endpoint, content).ConfigureAwait(true))
                {
                    var body = await response.Content.ReadAsStringAsync().ConfigureAwait(true);
                    var payload = new JavaScriptSerializer().DeserializeObject(string.IsNullOrWhiteSpace(body) ? "{}" : body)
                        as Dictionary<string, object> ?? new Dictionary<string, object>();
                    var token = payload.ContainsKey("token") ? Convert.ToString(payload["token"]) : string.Empty;
                    if (!response.IsSuccessStatusCode || string.IsNullOrWhiteSpace(token))
                    {
                        throw new InvalidOperationException(payload.ContainsKey("error") ? Convert.ToString(payload["error"]) : "云端登录失败。");
                    }

                    new CloudSessionStore().SaveToken(token);
                    _facade?.Dispose();
                    _facade = null;
                    Application.StatusBar = "校准助手：已登录 wzglpt.top";
                    _ = RefreshServiceConnectionStatusAsync();
                    return true;
                }
            }
        }

        internal void LogoutFromCloud()
        {
            new CloudSessionStore().Clear();
            _facade?.Dispose();
            _facade = null;
            Application.StatusBar = "校准助手：已退出云端登录";
        }

        internal CellRange GetActiveSelectionRange()
        {
            try
            {
                var selection = this.Application?.Selection as Excel.Range;
                if (selection == null)
                {
                    return null;
                }

                var worksheet = selection.Worksheet as Excel.Worksheet;
                if (worksheet == null)
                {
                    return null;
                }

                var startRow = selection.Row;
                var startColumn = selection.Column;
                return new CellRange
                {
                    SheetName = worksheet.Name,
                    StartRow = startRow,
                    StartColumn = startColumn,
                    EndRow = startRow + selection.Rows.Count - 1,
                    EndColumn = startColumn + selection.Columns.Count - 1
                };
            }
            catch
            {
                return null;
            }
        }

        internal void ShowRandomGenerationConfiguration()
        {
            var store = new GenerationConfigurationStore();
            var current = LoadGenerationConfiguration();
            using (var dialog = new RandomGenerationConfigurationDialog(current))
            {
                if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    return;
                }

                store.Save(dialog.Configuration);
                _generationConfiguration = store.Normalize(dialog.Configuration);
                _facade?.Dispose();
                _facade = null;
            }

            ApplyKeyboardShortcutConfiguration();
            ApplyGlobalGenerationConfigurationToCurrentState();
            UpdateCurrentGenerationRules(
                _taskPaneControl?.GetCurrentRules(),
                _taskPaneControl?.GetAppliedGenerationConfiguration() ?? _generationConfiguration);
            this.Application.StatusBar = "校准助手：随机数配置已保存";
            Trace.WriteLine("[VSTO] Random generation configuration saved.");
        }

        internal async Task<TemplateSaveResult> SaveCurrentTemplateAsync(
            string templateName,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration = null,
            bool createNew = false,
            TemplateDirectoryMetadata directoryMetadata = null)
        {
            EnsureFacade();
            var preparedRules = _facade.PrepareRulesForTemplateSave(this.Application?.ActiveWorkbook, rules);
            var result = await _facade.SaveTemplateAsync(
                templateName,
                fingerprint,
                preparedRules,
                generationConfiguration,
                createNew,
                directoryMetadata);
            this.Application.StatusBar = $"校准助手：{result.Message}";
            Trace.WriteLine($"[VSTO] Template saved. Name={templateName}, Rules={rules?.Count ?? 0}, SyncStatus={result.LocalSyncStatus}");
            return result;
        }

        internal TemplateSaveResult SaveCurrentTemplate(
            string templateName,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration = null,
            bool createNew = false,
            TemplateDirectoryMetadata directoryMetadata = null,
            bool prepareFromWorkbook = true,
            string targetRemoteTemplateId = null)
        {
            EnsureFacade();
            var preparedRules = prepareFromWorkbook
                ? _facade.PrepareRulesForTemplateSave(this.Application?.ActiveWorkbook, rules)
                : TaskPaneModelCloner.CloneRules(rules);
            var result = _facade.SaveTemplateAsync(
                    templateName,
                    fingerprint,
                    preparedRules,
                    generationConfiguration,
                    createNew,
                    directoryMetadata,
                    targetRemoteTemplateId)
                .GetAwaiter()
                .GetResult();
            this.Application.StatusBar = $"校准助手：{result.Message}";
            Trace.WriteLine($"[VSTO] Template saved. Name={templateName}, Rules={rules?.Count ?? 0}, SyncStatus={result.LocalSyncStatus}");
            return result;
        }

        internal IReadOnlyList<SavedTemplateInfo> GetSavedTemplates()
        {
            EnsureFacade();
            return _facade.ListSavedTemplates();
        }

        internal void ShowSaveSampleData()
        {
            EnsureFacade();
            var state = _lastGenerationState;
            if (state == null || !state.CanGenerate || string.IsNullOrWhiteSpace(state.ExactFingerprint)) throw new InvalidOperationException("请先识别并匹配模板。");
            using (var dialog = new SampleDataSelectionDialog(state.DraftRules))
            {
                if (dialog.ShowDialog(GetExcelMainWindow()) != DialogResult.OK || dialog.SelectedNames.Count == 0) return;
                var result = _facade.SaveSampleData(Application.ActiveWorkbook, state.ExactFingerprint, state.DraftRules, new HashSet<string>(dialog.SelectedNames, StringComparer.Ordinal));
                var message = $"已保存 {result.SavedItemCount} 个校准项。" + (result.SkippedItems.Count == 0 ? string.Empty : Environment.NewLine + "跳过：" + Environment.NewLine + string.Join(Environment.NewLine, result.SkippedItems));
                MessageBox.Show(message, "保存样本数据", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        internal void ShowSampleDataVersions()
        {
            EnsureFacade();
            var state = _lastGenerationState;
            if (state == null || string.IsNullOrWhiteSpace(state.ExactFingerprint)) throw new InvalidOperationException("请先识别并匹配模板。");
            using (var dialog = new SampleDataVersionDialog(_facade.ListSampleDataVersions(state.ExactFingerprint), id => _facade.DeleteSampleDataVersion(id))) dialog.ShowDialog(GetExcelMainWindow());
        }

        internal void ShowTemplateLibraryManager()
        {
            EnsureFacade();
            var shouldRefreshCurrentWorkbook = false;
            using (var dialog = new TemplateLibraryManagerDialog(
                () => _facade.ListSavedTemplates(),
                (exactFingerprint, status) => _facade.UpdateSavedTemplateStatus(exactFingerprint, status),
                exactFingerprint => _facade.DeleteSavedTemplate(exactFingerprint),
                exactFingerprint => _facade.GetSavedTemplateGenerationConfiguration(exactFingerprint),
                (exactFingerprint, configuration) =>
                {
                    var updated = _facade.UpdateSavedTemplateGenerationConfiguration(exactFingerprint, configuration);
                    shouldRefreshCurrentWorkbook = shouldRefreshCurrentWorkbook || updated;
                    return updated;
                },
                () => _facade.UploadPendingTemplates(),
                () => _facade.SyncTemplatesAsync().GetAwaiter().GetResult(),
                (exactFingerprint, action, saveAsTemplateName) =>
                {
                    var resolved = _facade.ResolveTemplateConflict(exactFingerprint, action, saveAsTemplateName);
                    shouldRefreshCurrentWorkbook = shouldRefreshCurrentWorkbook || resolved;
                    return resolved;
                },
                path => _facade.ExportDiagnostics(path, _lastGenerationState),
                LoadGenerationConfiguration,
                template => BeginEditSavedTemplate(template),
                exactFingerprint => _facade.ListSampleDataVersions(exactFingerprint),
                versionId => _facade.DeleteSampleDataVersion(versionId)))
            {
                dialog.ShowDialog();
            }

            if (shouldRefreshCurrentWorkbook)
            {
                _ = MatchWorkbookIfAvailableAsync(this.Application?.ActiveWorkbook as Excel.Workbook, force: true);
            }
        }

        private bool BeginEditSavedTemplate(SavedTemplateInfo template)
        {
            if (template == null)
            {
                return false;
            }

            EnsureTaskPane();
            if (!_taskPaneControl.ConfirmCloseIfDirty())
            {
                return false;
            }
            var state = _facade.BuildSavedTemplateEditorState(template.ExactFingerprint, template.RemoteTemplateId);
            if (state == null)
            {
                MessageBox.Show("模板不存在或已被删除。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            _taskPaneControl.SetReadOnlyMode(false);
            _taskPaneControl.BeginEditSavedTemplate(state, template);
            SetTaskPaneVisible(true);
            return true;
        }

        internal string GetActiveSelectionQualitySummary()
        {
            try
            {
                var selection = this.Application?.Selection as Excel.Range;
                if (selection == null)
                {
                    return string.Empty;
                }

                var rowCount = selection.Rows.Count;
                var columnCount = selection.Columns.Count;
                var hasFormula = Convert.ToString(selection.HasFormula);
                var hasMergedCells = Convert.ToBoolean(selection.MergeCells);
                var details = new List<string>
                {
                    $"行数 {rowCount}",
                    $"列数 {columnCount}"
                };

                if (string.Equals(hasFormula, "True", StringComparison.OrdinalIgnoreCase))
                {
                    details.Add("包含公式");
                }
                else if (string.Equals(hasFormula, "False", StringComparison.OrdinalIgnoreCase))
                {
                    details.Add("不包含公式");
                }
                else
                {
                    details.Add("部分单元格包含公式");
                }

                details.Add(hasMergedCells ? "包含合并单元格" : "不包含合并单元格");
                return string.Join("；", details);
            }
            catch
            {
                return string.Empty;
            }
        }

        internal async Task RecognizeCurrentWorkbookAsync()
        {
            EnsureFacade();
            EnsureTaskPane();
            if (!_taskPaneControl.ConfirmCloseIfDirty())
            {
                return;
            }

            var workbook = this.Application?.ActiveWorkbook;
            if (workbook == null)
            {
                Trace.WriteLine("[VSTO] Recognize skipped: no active workbook.");
                return;
            }

            var activeSheetName = string.Empty;
            try
            {
                activeSheetName = Convert.ToString(this.Application?.ActiveSheet?.Name);
            }
            catch
            {
                activeSheetName = string.Empty;
            }

            Trace.WriteLine($"[VSTO] Recognize start. Workbook={workbook.Name}, Sheet={activeSheetName}");

            _lastHighlightedRange = null;
            SetTaskPaneVisible(true);
            _taskPaneControl.SetRecognitionProgress("正在识别模板...", 0, true);
            this.Application.StatusBar = "校准助手：正在识别模板...";
            RecalculateWorkbook(workbook);

            try
            {
                var state = await _facade.RecognizeAsync(workbook, ReportRecognitionProgress);
                Trace.WriteLine(
                    $"[VSTO] Recognize finished. Workbook={state.WorkbookName}, " +
                    $"Fingerprint={state.ExactFingerprint}, Status={state.RemoteStatus}, " +
                    $"Mappings={state.MappingItems?.Count ?? 0}, Fields={state.RecognizedFields?.Count ?? 0}");
                await ApplyRecognitionActionAsync(workbook, state);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Recognize failed: {ex}");
                throw;
            }
            finally
            {
                _taskPaneControl.SetRecognitionProgress("识别完成", 100, false);
                this.Application.StatusBar = false;
            }
        }

        private async Task ApplyRecognitionActionAsync(Excel.Workbook workbook, TaskPaneState state)
        {
            var workbookKey = ResolveWorkbookKey(workbook);
            RememberGenerationState(workbookKey, state);

            if (state?.CanGenerate != true)
            {
                if (state?.IsCandidateMatch == true)
                {
                    await ApplyCandidateRecognitionActionAsync(workbook, workbookKey, state);
                    return;
                }

                _taskPaneControl.SetReadOnlyMode(false);
                _taskPaneControl.Bind(state);
                RefreshRandomRangeSummary(state?.DraftRules, state?.AppliedGenerationConfiguration);
                return;
            }

            var action = TemplateRecognitionActionDialog.ShowDialog(GetExcelMainWindow(), state.RemoteTemplateName);
            if (action == TemplateRecognitionAction.Close)
            {
                SetTaskPaneVisible(false);
                RefreshRandomRangeSummary(state.DraftRules, state.AppliedGenerationConfiguration);
                return;
            }

            if (action == TemplateRecognitionAction.View)
            {
                _taskPaneControl.SetReadOnlyMode(true);
                _taskPaneControl.Bind(state);
                RefreshRandomRangeSummary(state.DraftRules, state.AppliedGenerationConfiguration);
                return;
            }

            if (action == TemplateRecognitionAction.Edit)
            {
                _taskPaneControl.SetReadOnlyMode(false);
                _taskPaneControl.Bind(state);
                RefreshRandomRangeSummary(state.DraftRules, state.AppliedGenerationConfiguration);
                return;
            }

            var draftState = await _facade.RecognizeDraftAsync(workbook, ReportRecognitionProgress);
            RememberGenerationState(workbookKey, draftState);
            _taskPaneControl.SetReadOnlyMode(false);
            _taskPaneControl.Bind(draftState);
            _taskPaneControl.SetPreferredCreateNewTemplate(action == TemplateRecognitionAction.SaveAs);
            RefreshRandomRangeSummary(draftState.DraftRules, draftState.AppliedGenerationConfiguration);
        }

        private async Task ApplyCandidateRecognitionActionAsync(
            Excel.Workbook workbook,
            string workbookKey,
            TaskPaneState state)
        {
            var action = TemplateCandidateActionDialog.ShowDialog(
                GetExcelMainWindow(),
                state.RemoteTemplateName,
                state.LocalMatchScore);

            if (action == TemplateCandidateAction.Close)
            {
                SetTaskPaneVisible(false);
                return;
            }

            if (action == TemplateCandidateAction.View)
            {
                _taskPaneControl.SetReadOnlyMode(true);
                _taskPaneControl.Bind(state);
                RefreshRandomRangeSummary(state.DraftRules, state.AppliedGenerationConfiguration);
                return;
            }

            var draftState = await _facade.RecognizeDraftAsync(workbook, ReportRecognitionProgress);
            RememberGenerationState(workbookKey, draftState);
            _taskPaneControl.SetReadOnlyMode(false);
            _taskPaneControl.Bind(draftState);
            _taskPaneControl.SetPreferredCreateNewTemplate(action == TemplateCandidateAction.SaveAs);
            RefreshRandomRangeSummary(draftState.DraftRules, draftState.AppliedGenerationConfiguration);
        }

        internal void HighlightRange(CellRange range)
        {
            if (range == null || string.IsNullOrWhiteSpace(range.SheetName))
            {
                return;
            }

            if (IsSameRange(_lastHighlightedRange, range))
            {
                return;
            }

            try
            {
                _isNavigatingFromPlugin = true;
                var worksheet = this.Application?.ActiveWorkbook?.Worksheets[range.SheetName] as Excel.Worksheet;
                if (worksheet == null)
                {
                    return;
                }

                var target = worksheet.Range[worksheet.Cells[range.StartRow, range.StartColumn], worksheet.Cells[range.EndRow, range.EndColumn]];
                worksheet.Activate();
                if (IsRangeFullyVisible(target))
                {
                    target.Select();
                }
                else
                {
                    this.Application.Goto(target, true);
                    target.Select();
                }

            _lastHighlightedRange = TaskPaneModelCloner.CloneRange(range);
            }
            catch
            {
            }
            finally
            {
                _isNavigatingFromPlugin = false;
            }
        }

        private System.Windows.Forms.IWin32Window GetExcelMainWindow()
        {
            try
            {
                var handle = new IntPtr(this.Application.Hwnd);
                return handle == IntPtr.Zero ? null : new ExcelMainWindow(handle);
            }
            catch
            {
                return null;
            }
        }

        private sealed class ExcelMainWindow : System.Windows.Forms.IWin32Window
        {
            public ExcelMainWindow(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
        }

    }
}
