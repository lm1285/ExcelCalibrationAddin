using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Templates;
using Newtonsoft.Json.Linq;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public sealed partial class SyncTemplateUseCase
    {
        private const string LastSuccessfulSyncPreference = "templates.last_successful_sync_utc";

        public async Task<TemplateSyncRunResult> SyncAsync()
        {
            var result = new TemplateSyncRunResult();
            try
            {
                result.PendingUploadsSucceeded = UploadPendingTemplates();
                var raw = await _syncClient.ListTemplatesAsync();
                var root = JObject.Parse(raw);
                var templates = root["data"] as JArray ?? new JArray();
                foreach (var token in templates)
                {
                    var remote = ParseRemoteTemplate(token);
                    if (remote == null)
                    {
                        result.FailedCount++;
                        continue;
                    }

                    try
                    {
                        var applyResult = _cacheRepository.UpsertRemoteTemplate(
                            remote.TemplateId,
                            remote.TemplateName,
                            remote.Version,
                            remote.UpdatedAt,
                            remote.DeletedAt,
                            remote.Status,
                            remote.Fingerprint,
                            remote.Rules,
                            remote.GenerationConfiguration,
                            remote.DirectoryMetadata);
                        result.ProcessedCount++;
                        if (applyResult == ExcelCalibrationAddin.Core.Repositories.TemplateRemoteApplyResult.Applied)
                        {
                            result.AppliedCount++;
                        }
                        else if (applyResult == ExcelCalibrationAddin.Core.Repositories.TemplateRemoteApplyResult.Conflict)
                        {
                            result.ConflictCount++;
                        }
                        else
                        {
                            result.IgnoredCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedCount++;
                        Trace.WriteLine($"[Host] Remote template apply failed. Name={remote.TemplateName}, Error={ex.Message}");
                    }
                }

                _cacheRepository.SetPreference(LastSuccessfulSyncPreference, DateTime.UtcNow.ToString("o"));
                result.Succeeded = true;
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                Trace.WriteLine($"[Host] Template sync failed. Error={ex}");
                return result;
            }
        }

        public DateTime? GetLastSuccessfulSyncUtc()
        {
            DateTime value;
            var raw = _cacheRepository.GetPreference(LastSuccessfulSyncPreference);
            return DateTime.TryParse(raw, out value) ? value.ToUniversalTime() : (DateTime?)null;
        }

        private int UploadPendingTemplates()
        {
            var pending = _cacheRepository.ListTemplatesBySyncStatus(
                TemplateSyncStatus.PendingUpload,
                TemplateSyncStatus.SyncFailed);
            var uploaded = 0;
            foreach (var template in pending)
            {
                try
                {
                    var remoteResponse = _syncClient.SaveTemplateAsync(
                        template.TemplateName,
                        template.Fingerprint,
                        template.Rules,
                        template.GenerationConfiguration,
                        template.RemoteTemplateId,
                        IsLocallyCreatedTemplateId(template.RemoteTemplateId),
                        template.DirectoryMetadata).GetAwaiter().GetResult();
                    var acceptedTemplate = RemoteTemplateSaveResponseParser.Parse(
                        remoteResponse,
                        template.TemplateName,
                        template.Fingerprint,
                        template.Rules,
                        template.GenerationConfiguration);
                    _cacheRepository.SaveAcceptedRemoteTemplate(
                        acceptedTemplate.TemplateId,
                        acceptedTemplate.TemplateName,
                        acceptedTemplate.Version,
                        acceptedTemplate.RemoteUpdatedAt,
                        acceptedTemplate.DeletedAt,
                        acceptedTemplate.Status,
                        acceptedTemplate.Fingerprint,
                        acceptedTemplate.Rules,
                        acceptedTemplate.GenerationConfiguration,
                        template.DirectoryMetadata,
                        template.RemoteTemplateId);
                    uploaded++;
                }
                catch (Exception ex)
                {
                    _cacheRepository.UpdateTemplateSyncStatus(template.ExactFingerprint, TemplateSyncStatus.SyncFailed, ex.Message);
                }
            }

            return uploaded;
        }

        private static bool IsLocallyCreatedTemplateId(string templateId)
        {
            return string.IsNullOrWhiteSpace(templateId) ||
                templateId.StartsWith("template:", StringComparison.OrdinalIgnoreCase);
        }

        private static RemoteTemplateRecord ParseRemoteTemplate(JToken token)
        {
            if (token == null)
            {
                return null;
            }

            var fingerprint = ParseFingerprint(token["fingerprint_hash"] ?? token["fingerprint"]);
            if (fingerprint == null || string.IsNullOrWhiteSpace(fingerprint.ExactFingerprint))
            {
                return null;
            }

            var templateId = token["template_id"]?.Value<string>() ?? token["id"]?.Value<string>();
            var templateName = token["template_name"]?.Value<string>() ?? token["name"]?.Value<string>();
            if (string.IsNullOrWhiteSpace(templateId) || string.IsNullOrWhiteSpace(templateName))
            {
                return null;
            }

            var deletedAt = ParseDate(token["deleted_at"]);
            return new RemoteTemplateRecord
            {
                TemplateId = templateId,
                TemplateName = templateName,
                Version = token["version"]?.Value<int?>() ?? token["remote_version"]?.Value<int?>() ?? 0,
                UpdatedAt = ParseDate(token["updated_at"]),
                DeletedAt = deletedAt,
                Status = deletedAt.HasValue ? TemplateLifecycleStatus.Obsolete : ParseStatus(token["status"]) ?? TemplateLifecycleStatus.Enabled,
                Fingerprint = fingerprint,
                Rules = ParseRules(token["rules_json"] ?? token["rules"]),
                GenerationConfiguration = ParseGenerationConfiguration(token["generation_config_json"] ?? token["generationConfiguration"]),
                DirectoryMetadata = ParseDirectoryMetadata(token["directory_metadata"] ?? token["directoryMetadata"])
            };
        }

        private static TemplateFingerprint ParseFingerprint(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            try
            {
                return token.Type == JTokenType.String
                    ? JObject.Parse(token.Value<string>()).ToObject<TemplateFingerprint>()
                    : token.ToObject<TemplateFingerprint>();
            }
            catch
            {
                return null;
            }
        }

        private static TemplateDirectoryMetadata ParseDirectoryMetadata(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            try
            {
                return token.Type == JTokenType.String
                    ? JObject.Parse(token.Value<string>()).ToObject<TemplateDirectoryMetadata>()
                    : token.ToObject<TemplateDirectoryMetadata>();
            }
            catch
            {
                return null;
            }
        }

        private static DateTime? ParseDate(JToken token)
        {
            DateTime value;
            return DateTime.TryParse(token?.ToString(), out value) ? value.ToUniversalTime() : (DateTime?)null;
        }
    }

    public sealed class TemplateSyncRunResult
    {
        public bool Succeeded { get; set; }
        public int ProcessedCount { get; set; }
        public int AppliedCount { get; set; }
        public int IgnoredCount { get; set; }
        public int ConflictCount { get; set; }
        public int FailedCount { get; set; }
        public int PendingUploadsSucceeded { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
    }

    internal sealed class RemoteTemplateRecord
    {
        public string TemplateId { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public int Version { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public DateTime? DeletedAt { get; set; }
        public TemplateLifecycleStatus Status { get; set; }
        public TemplateFingerprint Fingerprint { get; set; }
        public IReadOnlyList<MeasurementRule> Rules { get; set; } = new List<MeasurementRule>();
        public ExcelCalibrationAddin.Core.Models.GenerationConfiguration GenerationConfiguration { get; set; }
        public TemplateDirectoryMetadata DirectoryMetadata { get; set; }
    }
}
