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
        private bool ShouldBlockFeatures(RecognitionAndDraftResult result)
        {
            if (IsDevelopmentBackend())
            {
                return HasInactiveLocalTemplate(result);
            }

            return HasBackendFailure(result) ||
                HasFingerprintFailure(result) ||
                HasInactiveLocalTemplate(result) ||
                HasInactiveRemoteTemplate(result);
        }

        private static bool HasBackendFailure(RecognitionAndDraftResult result)
        {
            return !string.IsNullOrWhiteSpace(result?.Remote?.ErrorMessage);
        }

        private static bool HasFingerprintFailure(RecognitionAndDraftResult result)
        {
            return string.IsNullOrWhiteSpace(result?.Recognition?.Fingerprint?.ExactFingerprint);
        }

        private static bool HasInactiveLocalTemplate(RecognitionAndDraftResult result)
        {
            return result?.Local != null &&
                result.Local.Status != TemplateLifecycleStatus.Enabled;
        }

        private static bool HasInactiveRemoteTemplate(RecognitionAndDraftResult result)
        {
            return result?.Remote?.Status.HasValue == true &&
                result.Remote.Status.Value != TemplateLifecycleStatus.Enabled;
        }

        private bool IsDevelopmentBackend()
        {
            var baseUrl = _bootstrapper.Configuration?.Backend?.BaseUrl ?? string.Empty;
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                return false;
            }

            try
            {
                var uri = new Uri(baseUrl);
                return string.Equals(uri.Host, "localhost", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(uri.Host, "127.0.0.1", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(uri.Host, "::1", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return baseUrl.IndexOf("localhost", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    baseUrl.IndexOf("127.0.0.1", StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        private static string BuildRemoteDetail(RecognitionAndDraftResult result)
        {
            if (!string.IsNullOrWhiteSpace(result.Remote.ErrorMessage))
            {
                return $"远端模板库不可用: {result.Remote.ErrorMessage}";
            }

            if (result.Remote.Found)
            {
                var status = result.Remote.Status.HasValue
                    ? TranslateTemplateStatus(result.Remote.Status.Value)
                    : "未知";
                var name = string.IsNullOrWhiteSpace(result.Remote.TemplateName) ? "未命名模板" : result.Remote.TemplateName;
                return $"命中远端模板 {name}, 分数 {result.Remote.MatchScore:F0}, 状态 {status}";
            }

            if (result.Local != null)
            {
                var detail = string.IsNullOrWhiteSpace(result.Local.TemplateName)
                    ? "命中本地模板"
                    : $"命中本地模板 {result.Local.TemplateName}";

                if (result.Local.Status == TemplateLifecycleStatus.Obsolete ||
                    result.Local.Status == TemplateLifecycleStatus.Disabled)
                {
                    return $"{detail} | 状态: {TranslateTemplateStatus(result.Local.Status)}";
                }

                if (result.Local.MatchScore < 100)
                {
                    return $"{detail} | 候选匹配, 需要用户确认后另存或重新识别";
                }

                return $"{detail} | 状态: 启用";
            }

            return result.Recognition.RecognizedFields.Count > 0
                ? "已提取模板字段, 但模板库中未找到足够接近的模板"
                : "未提取到可用于匹配的模板字段";
        }
        private static string ResolveTemplateId(RecognitionAndDraftResult result)
        {
            if (!string.IsNullOrWhiteSpace(result.Remote.TemplateId))
            {
                return result.Remote.TemplateId;
            }

            return result.Local?.RemoteTemplateId ?? string.Empty;
        }

        private static string ResolveTemplateName(RecognitionAndDraftResult result)
        {
            if (!string.IsNullOrWhiteSpace(result.Remote.TemplateName))
            {
                return result.Remote.TemplateName;
            }

            return result.Local?.TemplateName ?? string.Empty;
        }

        private GenerationConfiguration ResolveAppliedGenerationConfiguration(ExcelCalibrationAddin.Core.Repositories.CachedTemplateRule localTemplate)
        {
            var store = new ExcelCalibrationAddin.Core.Services.GenerationConfigurationStore();
            return localTemplate?.GenerationConfiguration != null
                ? store.Clone(localTemplate.GenerationConfiguration)
                : store.Clone(_bootstrapper.Configuration?.Generation);
        }

        private GenerationConfiguration ResolveAppliedGenerationConfiguration(RecognitionAndDraftResult result)
        {
            var store = new ExcelCalibrationAddin.Core.Services.GenerationConfigurationStore();
            if (result?.Local?.GenerationConfiguration != null)
            {
                return store.Clone(result.Local.GenerationConfiguration);
            }

            if (result?.Remote?.GenerationConfiguration != null)
            {
                return store.Clone(result.Remote.GenerationConfiguration);
            }

            return store.Clone(_bootstrapper.Configuration?.Generation);
        }

        private static bool HasTemplateGenerationConfiguration(RecognitionAndDraftResult result)
        {
            return result?.Local?.GenerationConfiguration != null ||
                result?.Remote?.GenerationConfiguration != null;
        }

        private static bool HasMatchedGenerationRules(RecognitionAndDraftResult result)
        {
            return (result?.Local?.Rules != null &&
                    result.Local.Rules.Count > 0 &&
                    result.Local.Status == TemplateLifecycleStatus.Enabled &&
                    result.Local.MatchScore >= 100) ||
                (result?.Remote?.Rules != null &&
                 result.Remote.Rules.Count > 0 &&
                 (!result.Remote.Status.HasValue || result.Remote.Status.Value == TemplateLifecycleStatus.Enabled));
        }

        private static string TranslateTemplateStatus(TemplateLifecycleStatus status)
        {
            switch (status)
            {
                case TemplateLifecycleStatus.Disabled:
                    return "停用";
                case TemplateLifecycleStatus.Obsolete:
                    return "废止";
                default:
                    return "启用";
            }
        }

        private static string TranslateSyncStatus(TemplateSyncStatus status)
        {
            switch (status)
            {
                case TemplateSyncStatus.PendingUpload:
                    return "待上传";
                case TemplateSyncStatus.Conflict:
                    return "冲突";
                case TemplateSyncStatus.SyncFailed:
                    return "同步失败";
                default:
                    return "已同步";
            }
        }
        private GenerationConfiguration ResolveCurrentGenerationConfiguration()
        {
            var store = new ExcelCalibrationAddin.Core.Services.GenerationConfigurationStore();
            return store.Clone(_bootstrapper.Configuration?.Generation);
        }

    }
}
