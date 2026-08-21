using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Host.Generation;
using ExcelCalibrationAddin.Host.Vsto;
using ExcelCalibrationAddin.Host.ViewModels;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private const int GenerationRetryLimit = 100;

        internal async Task GenerateRandomNumbersCurrentWorkbookAsync()
        {
            EnsureFacade();

            var workbook = this.Application?.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("请先打开一个工作 Excel。");
            }

            var totalWatch = Stopwatch.StartNew();
            WriteGenerationPerformanceLog($"start version={AddinVersion.Version} workbook={workbook.Name}");
            Trace.WriteLine($"[VSTO] GenerateRandom start. Version={AddinVersion.Version}, Workbook={workbook.Name}");

            _lastHighlightedRange = null;

            var calculationWasAuto = false;
            try
            {
                var generationOverride = Globals.Ribbons?.CalibrationRibbon?.GetSingleUseOverride();
                var workbookKey = ResolveWorkbookKey(workbook);
                if (!CanUseCachedGenerationState(workbookKey) && !HasCachedWorkbookMatchState(workbookKey))
                {
                    await MatchWorkbookIfAvailableAsync(workbook, force: false);
                }

                if (!CanUseCachedGenerationState(workbookKey))
                {
                    throw new InvalidOperationException(
                        "当前工作簿未匹配到已启用的可生成模板。请确认模板已保存并启用。");
                }

                this.Application.StatusBar = "校准助手：正在生成随机数...";
                System.Windows.Forms.Application.DoEvents();
                using (var performanceScope = EnterFastGenerationMode())
                using (var transaction = new ExcelGenerationTransaction(workbook, _lastGenerationState.DraftRules))
                {
                    calculationWasAuto = performanceScope.CalculationWasAutomatic;
                    TaskPaneState state = null;
                    for (var attempt = 1; attempt <= GenerationRetryLimit; attempt++)
                    {
                        try
                        {
                            state = TryGenerateFromCachedState(workbook, workbookKey, generationOverride);
                            if (state == null)
                            {
                                throw new InvalidOperationException("已匹配模板未包含可生成的规则。");
                            }

                            RecalculateWorkbook(workbook);
                            _facade.VerifyFormulaResults(workbook, state.DraftRules);
                            break;
                        }
                        catch (GenerationRetryException exception) when (attempt < GenerationRetryLimit)
                        {
                            this.Application.StatusBar =
                                $"校准助手：误差结果为 0，正在自动重算（{attempt}/{GenerationRetryLimit}）...";
                            Trace.WriteLine(
                                $"[VSTO] Generation retry {attempt}/{GenerationRetryLimit}. {exception.Message}");
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }

                    Globals.Ribbons?.CalibrationRibbon?.ClearSingleUseOverride();
                    transaction.Commit();
                    RememberGenerationState(workbookKey, state);
                    RefreshRandomRangeSummary(state.DraftRules, state.AppliedGenerationConfiguration);
                    ShowGenerationWarnings(state.GenerationWarningMessages);
                }

                this.Application.StatusBar = "校准助手：随机数生成完成";
                WriteGenerationPerformanceLog($"finished elapsed={totalWatch.ElapsedMilliseconds}ms");
                Trace.WriteLine($"[VSTO] GenerateRandom finished. Workbook={workbook.Name}, elapsed={totalWatch.ElapsedMilliseconds}ms");
            }
            finally
            {
                WriteGenerationPerformanceLog($"exit elapsed={totalWatch.ElapsedMilliseconds}ms");
                Trace.WriteLine($"[VSTO] GenerateRandom exit elapsed={totalWatch.ElapsedMilliseconds}ms");
            }

        }

        private void ShowGenerationWarnings(IReadOnlyList<string> warnings)
        {
            var warningList = (warnings ?? new List<string>())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (warningList.Count == 0)
            {
                return;
            }

            var message = "生成过程中存在以下提示：" + Environment.NewLine + string.Join(Environment.NewLine, warningList);
            this.Application.StatusBar = "校准助手：生成完成，部分项目存在提示";
            System.Windows.Forms.MessageBox.Show(
                message,
                "校准助手",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

    }
}
