using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Host.Controllers;
using ExcelCalibrationAddin.Host.Interop;
using ExcelCalibrationAddin.Host.UseCases;
using ExcelCalibrationAddin.Host.ViewModels;

namespace ExcelCalibrationAddin.Host.Vsto
{
    public sealed partial class VstoAddinFacade
    {
        private AddinWorkflowController CreateController(dynamic workbook, string requestedSheetName = null)
        {
            var snapshotProvider = new ExcelInteropSnapshotProvider(workbook, requestedSheetName);
            var writer = new ExcelInteropWriter(workbook);
            var orchestrator = _bootstrapper.CreateWorkflowOrchestrator(snapshotProvider, writer);
            return new AddinWorkflowController(
                orchestrator,
                _bootstrapper.MeasurementRuleDraftBuilder,
                _bootstrapper.MeasurementRuleParameterResolver);
        }

        private TaskPaneState BuildTaskPaneState(
            RecognitionAndDraftResult result,
            IReadOnlyList<MeasurementRule> draftRules,
            bool canGenerate,
            GenerationConfiguration appliedGenerationConfiguration = null,
            bool usesTemplateGenerationConfiguration = false)
        {
            var resolvedConfiguration = appliedGenerationConfiguration ??
                ResolveAppliedGenerationConfiguration(result);

            return new TaskPaneState
            {
                WorkbookName = result.Recognition.Snapshot.WorkbookName,
                ExactFingerprint = result.Recognition.Fingerprint.ExactFingerprint,
                RemoteTemplateId = ResolveTemplateId(result),
                RemoteTemplateName = ResolveTemplateName(result),
                RemoteStatus = ResolveRemoteStatus(result, canGenerate),
                RemoteDetail = BuildRemoteDetail(result),
                MatchStatus = BuildMatchStatus(result),
                MatchStatusDetail = BuildMatchStatusDetail(result),
                RecognitionStatusDetail = BuildRecognitionStatusDetail(result),
                FingerprintStatusDetail = BuildFingerprintStatusDetail(result),
                LocalMatchStatusDetail = BuildLocalMatchStatusDetail(result),
                RemoteMatchStatusDetail = BuildRemoteMatchStatusDetail(result),
                DraftRuleStatusDetail = BuildDraftRuleStatusDetail(result, draftRules, result.Mappings),
                LocalTemplateStatus = result.Local?.Status,
                LocalTemplateStatusDetail = BuildLocalTemplateStatusDetail(result),
                LocalMatchScore = result.Local?.MatchScore ?? 0,
                IsCandidateMatch = result.Local != null &&
                    result.Local.Status == TemplateLifecycleStatus.Enabled &&
                    result.Local.MatchScore >= 60 &&
                    result.Local.MatchScore < 100,
                AppliedGenerationConfiguration = resolvedConfiguration,
                UsesTemplateGenerationConfiguration = usesTemplateGenerationConfiguration || HasTemplateGenerationConfiguration(result),
                Fingerprint = result.Recognition.Fingerprint,
                IsFeatureBlocked = ShouldBlockFeatures(result),
                CanGenerate = canGenerate,
                RecognizedFields = result.Recognition.RecognizedFields,
                MappingItems = result.Mappings,
                DraftRules = draftRules ?? new List<MeasurementRule>()
            };
        }

        private static string ResolveRemoteStatus(RecognitionAndDraftResult result, bool canGenerate)
        {
            if (!string.IsNullOrWhiteSpace(result.Remote.ErrorMessage))
            {
                return "RemoteUnavailable";
            }

            if (result.Remote.Found)
            {
                return "Matched";
            }

            if (result.Local != null)
            {
                return canGenerate ? "Matched" : "DraftReady";
            }

            return result.Recognition.RecognizedFields.Count > 0 ? "DraftReady" : "Unrecognized";
        }

        private string BuildMatchStatus(RecognitionAndDraftResult result)
        {
            if (HasInactiveLocalTemplate(result))
            {
                return "Failed";
            }

            return ShouldBlockFeatures(result) ? "Failed" : "Normal";
        }

        private string BuildMatchStatusDetail(RecognitionAndDraftResult result)
        {
            if (HasInactiveLocalTemplate(result))
            {
                var name = ResolveTemplateName(result);
                var status = TranslateTemplateStatus(result.Local.Status);
                return string.IsNullOrWhiteSpace(name) ? $"模板已{status}" : $"模板已{status}: {name}";
            }

            if (HasBackendFailure(result))
            {
                return $"后端失败: {result.Remote.ErrorMessage}";
            }

            if (HasFingerprintFailure(result))
            {
                return "指纹失败: 未生成有效模板指纹。";
            }

            if (IsDevelopmentBackend())
            {
                return "正常: 开发环境已跳过后端阻塞检查";
            }

            return "正常";
        }

        private static string BuildLocalTemplateStatusDetail(RecognitionAndDraftResult result)
        {
            if (result?.Local == null)
            {
                return string.Empty;
            }

            return $"本地模板状态: {TranslateTemplateStatus(result.Local.Status)}; 同步状态: {TranslateSyncStatus(result.Local.LocalSyncStatus)}";
        }

        private static string BuildRecognitionStatusDetail(RecognitionAndDraftResult result)
        {
            var fieldCount = result?.Recognition?.RecognizedFields?.Count ?? 0;
            return fieldCount > 0 ? $"通过: 识别到 {fieldCount} 个项目区域" : "异常: 未识别到项目区域";
        }

        private static string BuildFingerprintStatusDetail(RecognitionAndDraftResult result)
        {
            return string.IsNullOrWhiteSpace(result?.Recognition?.Fingerprint?.ExactFingerprint)
                ? "异常: 未生成有效指纹"
                : "通过: 已生成模板指纹";
        }

        private static string BuildLocalMatchStatusDetail(RecognitionAndDraftResult result)
        {
            if (result?.Local == null)
            {
                return "未命中: 本地模板库无精确匹配";
            }

            var name = string.IsNullOrWhiteSpace(result.Local.TemplateName) ? "未命名模板" : result.Local.TemplateName;
            if (result.Local.Status != TemplateLifecycleStatus.Enabled)
            {
                return $"异常: 命中已{TranslateTemplateStatus(result.Local.Status)}本地模板 {name}";
            }

            if (result.Local.LocalSyncStatus != TemplateSyncStatus.Synced)
            {
                return $"提示: 命中本地模板 {name}, 同步状态为 {TranslateSyncStatus(result.Local.LocalSyncStatus)}";
            }

            if (result.Local.MatchScore < 100)
            {
                return $"待确认: 发现相似本地模板 {name}, 不能直接用于生成";
            }

            return $"通过: 命中本地模板 {name}";
        }

        private static string BuildRemoteMatchStatusDetail(RecognitionAndDraftResult result)
        {
            if (!string.IsNullOrWhiteSpace(result?.Remote?.ErrorMessage))
            {
                return $"异常: 远端模板库不可用, {result.Remote.ErrorMessage}";
            }

            if (result?.Remote?.Found == true)
            {
                var name = string.IsNullOrWhiteSpace(result.Remote.TemplateName) ? "未命名模板" : result.Remote.TemplateName;
                var score = result.Remote.MatchScore > 0 ? $", 分数 {result.Remote.MatchScore:F0}" : string.Empty;
                return $"通过: 命中远端模板 {name}{score}";
            }

            return "未命中: 远端模板库无精确匹配";
        }

        private static string BuildDraftRuleStatusDetail(
            RecognitionAndDraftResult result,
            IReadOnlyList<MeasurementRule> draftRules,
            IReadOnlyList<TemplateRegionMapping> mappings)
        {
            var ruleCount = draftRules?.Count ?? 0;
            if (ruleCount == 0)
            {
                return "异常: 未生成可用规则";
            }

            if (result?.Local != null || result?.Remote?.Found == true)
            {
                return $"通过: 已加载 {ruleCount} 条模板规则";
            }

            var mappingCount = mappings?.Count ?? 0;
            if (mappingCount > ruleCount)
            {
                var missingMeasurementCount = mappings.Count(item => item?.MeasurementValueRange == null);
                var reason = missingMeasurementCount > 0
                    ? $"{missingMeasurementCount} 个区域未识别测量值"
                    : $"{mappingCount - ruleCount} 个区域未生成规则";
                return $"待确认: {ruleCount}/{mappingCount} 条可用, {reason}";
            }

            return $"待确认: 已生成 {ruleCount} 条规则";
        }
    }
}
