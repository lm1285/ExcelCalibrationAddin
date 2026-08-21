using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.UseCases;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            EnsureExcelUiSynchronizationContext();
            AddinFileLogger.Configure("VSTO");
            _configPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
            InitializeKeyboardShortcut();
            Application.WorkbookOpen += Application_WorkbookOpen;
            Application.WorkbookActivate += Application_WorkbookActivate;
            Application.SheetSelectionChange += Application_SheetSelectionChange;
            Application.SheetChange += Application_SheetChange;
            RefreshRandomRangeSummary(null);
            _dailySyncScheduler = new DailySyncScheduler(
                () => _facade == null ? RunTemplateSyncAsync() : _facade.SyncTemplatesAsync(),
                () =>
                {
                    EnsureFacade();
                    return _facade.GetLastSuccessfulSyncUtc();
                });
            _facadeWarmupTask = Task.Run((Action)EnsureFacade);
            _startupTemplateSyncTask = RunDeferredDailySyncAsync();
            _ = MatchWorkbookAfterOpenAsync(Application?.ActiveWorkbook as Excel.Workbook, force: false);
            Trace.WriteLine("[VSTO] Add-in startup completed.");
        }

        private async Task<TemplateSyncRunResult> RunTemplateSyncAsync()
        {
            EnsureFacade();
            return await _facade.SyncTemplatesAsync();
        }

        private async Task RunDeferredDailySyncAsync()
        {
            await Task.Delay(1500);
            var result = await Task.Run(() => _dailySyncScheduler.RunIfDueAsync());
            if (result == null)
            {
                return;
            }

            if (!result.Succeeded)
            {
                TrySetApplicationStatusBar("校准助手：模板后台同步失败，本地模板仍可继续使用");
                return;
            }

            if (result.FailedCount > 0 || result.ConflictCount > 0)
            {
                TrySetApplicationStatusBar($"校准助手：模板同步完成，失败 {result.FailedCount} 个，冲突 {result.ConflictCount} 个");
            }

            var workbook = Application?.ActiveWorkbook as Excel.Workbook;
            var workbookKey = ResolveWorkbookKey(workbook);
            if (workbook != null && !CanUseCachedGenerationState(workbookKey))
            {
                await MatchWorkbookIfAvailableAsync(workbook, force: true);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            DisposeKeyboardShortcut();
            if (_multiAreaCopyPasteDialog != null && !_multiAreaCopyPasteDialog.IsDisposed)
            {
                _multiAreaCopyPasteDialog.Close();
            }
            Application.WorkbookOpen -= Application_WorkbookOpen;
            Application.WorkbookActivate -= Application_WorkbookActivate;
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
            Application.SheetChange -= Application_SheetChange;
            _facade?.Dispose();
            _facade = null;
            Trace.WriteLine("[VSTO] Add-in shutdown completed.");
        }

        private async void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            Trace.WriteLine($"[VSTO] Workbook opened: {workbook?.Name}");
            _lastHighlightedRange = null;
            RefreshRandomRangeSummary(null, LoadGenerationConfiguration());
            await MatchWorkbookAfterOpenAsync(workbook, force: false);
        }

        private async void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            ApplyTaskPaneWidth();
            await MatchWorkbookAfterOpenAsync(workbook, force: false);
        }

        private async Task MatchWorkbookAfterOpenAsync(Excel.Workbook workbook, bool force)
        {
            if (workbook == null)
            {
                return;
            }

            await Task.Delay(100);
            await MatchWorkbookIfAvailableAsync(workbook, force);
        }

        private async Task MatchWorkbookIfAvailableAsync(Excel.Workbook workbook, bool force)
        {
            if (workbook == null)
            {
                return;
            }

            var workbookKey = ResolveWorkbookKey(workbook);
            Task matchTask;
            string matchedTaskWorkbookKey;
            lock (_workbookMatchTaskLock)
            {
                if (!force && HasCachedWorkbookMatchState(workbookKey))
                {
                    return;
                }

                if (!_activeWorkbookMatchTask.IsCompleted)
                {
                    matchTask = _activeWorkbookMatchTask;
                    matchedTaskWorkbookKey = _activeWorkbookMatchKey;
                }
                else
                {
                    _activeWorkbookMatchKey = workbookKey;
                    _activeWorkbookMatchTask = MatchWorkbookCoreAsync(workbook, workbookKey);
                    matchTask = _activeWorkbookMatchTask;
                    matchedTaskWorkbookKey = workbookKey;
                }
            }

            await matchTask;
            if (!string.Equals(matchedTaskWorkbookKey, workbookKey, StringComparison.OrdinalIgnoreCase))
            {
                await MatchWorkbookIfAvailableAsync(workbook, force);
            }
        }

        private async Task MatchWorkbookCoreAsync(Excel.Workbook workbook, string workbookKey)
        {
            await Task.Yield();
            try
            {
                await _facadeWarmupTask;
                EnsureFacade();
                TrySetApplicationStatusBar("校准助手：正在匹配模板库...");
                Trace.WriteLine($"[VSTO] Auto match start. Workbook={workbook.Name}");
                var state = await _facade.MatchTemplateLibraryAsync(workbook);
                RememberGenerationState(workbookKey, state);
                Globals.Ribbons?.CalibrationRibbon?.UpdateSampleDataButtonState();
                RefreshRandomRangeSummary(state.DraftRules, state.AppliedGenerationConfiguration);
                if (IsTaskPaneCurrentlyVisible() && _taskPaneControl != null &&
                    !_taskPaneControl.IsEditingSavedTemplate && !_taskPaneControl.HasUnsavedChanges)
                {
                    _taskPaneControl.Bind(state);
                }

                TrySetApplicationStatusBar(state.CanGenerate
                    ? "校准助手：已匹配模板库，可直接生成随机数"
                    : "校准助手：未匹配到可用模板，可先识别并保存模板");
                Trace.WriteLine(
                    $"[VSTO] Auto match finished. Workbook={state.WorkbookName}, " +
                    $"CanGenerate={state.CanGenerate}, Rules={state.DraftRules?.Count ?? 0}");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Auto match failed: {ex}");
                RefreshRandomRangeSummary(null, LoadGenerationConfiguration());
            }
            finally
            {
                TrySetApplicationStatusBar(false);
            }
        }

        private static void EnsureExcelUiSynchronizationContext()
        {
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(
                    new System.Windows.Forms.WindowsFormsSynchronizationContext());
            }
        }

        private void TrySetApplicationStatusBar(object value)
        {
            try
            {
                Application.StatusBar = value;
            }
            catch (COMException ex)
            {
                Trace.WriteLine($"[VSTO] Status bar update skipped: {ex.Message}");
            }
        }

        private void Application_SheetSelectionChange(object sh, Excel.Range target)
        {
            if (_isNavigatingFromPlugin || !IsTaskPaneCurrentlyVisible())
            {
                return;
            }
        }

        private void Application_SheetChange(object sh, Excel.Range target)
        {
            if (_isWritingGeneratedValues)
            {
                return;
            }

            if (IsChangeWithinCachedValueRanges(sh, target))
            {
                Trace.WriteLine("[VSTO] SheetChange retained cached template match for a dynamic value range.");
                return;
            }

            _lastGenerationState = null;
            _lastMatchedWorkbookKey = string.Empty;
            Globals.Ribbons?.CalibrationRibbon?.UpdateSampleDataButtonState();
        }

        private bool IsChangeWithinCachedValueRanges(object sheet, Excel.Range target)
        {
            if (_lastGenerationState?.DraftRules == null || target == null)
            {
                return false;
            }

            var sheetName = ResolveWorksheetName(sheet);
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return false;
            }

            var changedRange = ResolveExcelRange(target);
            if (changedRange == null)
            {
                return false;
            }

            return _lastGenerationState.DraftRules
                .Where(rule => rule != null)
                .SelectMany(GetDynamicValueRanges)
                .Where(range => range != null)
                .Any(range =>
                    string.Equals(range.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) &&
                    RangeContains(range, changedRange));
        }

        private static IEnumerable<CellRange> GetDynamicValueRanges(MeasurementRule rule)
        {
            yield return rule.TargetRange;
            yield return rule.SetpointSource?.Range;
            yield return rule.StandardValueSource?.Range;
            yield return rule.MpeSource?.Range;
            yield return rule.RangeSource?.Range;

            foreach (var mapping in rule.RowMappings ?? new List<MeasurementRowMapping>())
            {
                yield return mapping?.SetpointValueRange;
                yield return mapping?.StandardValueRange;
                yield return mapping?.TechnicalRequirementRange;
                yield return mapping?.RangeValueRange;
            }
        }

        private static string ResolveWorksheetName(object sheet)
        {
            try
            {
                return Convert.ToString(((Excel.Worksheet)sheet).Name) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static CellRange ResolveExcelRange(Excel.Range range)
        {
            try
            {
                return new CellRange
                {
                    StartRow = range.Row,
                    StartColumn = range.Column,
                    EndRow = range.Row + range.Rows.Count - 1,
                    EndColumn = range.Column + range.Columns.Count - 1
                };
            }
            catch
            {
                return null;
            }
        }

        private static bool RangeContains(CellRange outer, CellRange inner)
        {
            return outer != null &&
                inner != null &&
                inner.StartRow >= outer.StartRow &&
                inner.EndRow <= outer.EndRow &&
                inner.StartColumn >= outer.StartColumn &&
                inner.EndColumn <= outer.EndColumn;
        }
    }
}
