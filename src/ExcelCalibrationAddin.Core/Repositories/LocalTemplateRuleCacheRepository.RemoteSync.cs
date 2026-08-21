using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Repositories
{
    public sealed partial class LocalTemplateRuleCacheRepository
    {
        public TemplateRemoteApplyResult UpsertRemoteTemplate(
            string templateId,
            string templateName,
            int version,
            DateTime? remoteUpdatedAt,
            DateTime? deletedAt,
            TemplateLifecycleStatus status,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration,
            TemplateDirectoryMetadata directoryMetadata = null)
        {
            if (string.IsNullOrWhiteSpace(templateId) ||
                string.IsNullOrWhiteSpace(templateName) ||
                string.IsNullOrWhiteSpace(fingerprint?.ExactFingerprint))
            {
                throw new InvalidOperationException("远端模板缺少 ID、名称或指纹，无法同步。");
            }

            var templates = ListSavedTemplates();
            var existing = templates.FirstOrDefault(item =>
                string.Equals(item.RemoteTemplateId, templateId, StringComparison.OrdinalIgnoreCase));
            var sameName = templates
                .FirstOrDefault(item => string.Equals(item.TemplateName, templateName, StringComparison.OrdinalIgnoreCase));
            if (sameName != null &&
                (existing == null || !string.Equals(sameName.ExactFingerprint, existing.ExactFingerprint, StringComparison.OrdinalIgnoreCase)))
            {
                existing = sameName;
            }

            if (existing != null &&
                (existing.LocalSyncStatus == TemplateSyncStatus.PendingUpload ||
                 existing.LocalSyncStatus == TemplateSyncStatus.Conflict))
            {
                SaveRemoteConflictSnapshot(
                    existing.ExactFingerprint,
                    BuildConflictSnapshot(
                        templateId,
                        templateName,
                        version,
                        remoteUpdatedAt,
                        deletedAt,
                        status,
                        fingerprint,
                        rules,
                        generationConfiguration));
                return TemplateRemoteApplyResult.Conflict;
            }

            if (existing != null && IsOlderRemoteVersion(existing, version, remoteUpdatedAt))
            {
                return TemplateRemoteApplyResult.Ignored;
            }

            SaveRemoteTemplate(templateId, templateName, version, remoteUpdatedAt, deletedAt, status, fingerprint, rules, generationConfiguration, directoryMetadata);
            return TemplateRemoteApplyResult.Applied;
        }

        public void SaveAcceptedRemoteTemplate(
            string templateId,
            string templateName,
            int version,
            DateTime? remoteUpdatedAt,
            DateTime? deletedAt,
            TemplateLifecycleStatus status,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration,
            TemplateDirectoryMetadata directoryMetadata = null,
            string sourceRemoteTemplateId = null)
        {
            if (string.IsNullOrWhiteSpace(templateId) ||
                string.IsNullOrWhiteSpace(templateName) ||
                string.IsNullOrWhiteSpace(fingerprint?.ExactFingerprint))
            {
                throw new InvalidOperationException("远端模板缺少 ID、名称或指纹，无法写入本地缓存。");
            }

            SaveRemoteTemplate(templateId, templateName, version, remoteUpdatedAt, deletedAt, status, fingerprint, rules, generationConfiguration, directoryMetadata, sourceRemoteTemplateId);
        }

        private void SaveRemoteTemplate(
            string templateId,
            string templateName,
            int version,
            DateTime? remoteUpdatedAt,
            DateTime? deletedAt,
            TemplateLifecycleStatus status,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration,
            TemplateDirectoryMetadata directoryMetadata = null,
            string sourceRemoteTemplateId = null)
        {
            var now = DateTime.UtcNow.ToString("o");
            var remoteTime = remoteUpdatedAt?.ToUniversalTime().ToString("o") ?? string.Empty;
            var deletedTime = deletedAt?.ToUniversalTime().ToString("o") ?? string.Empty;
            var safeRules = (rules ?? new List<MeasurementRule>()).Where(item => item != null).ToList();
            var normalizedDirectoryMetadata = NormalizeDirectoryMetadata(directoryMetadata, templateName);

            using (var connection = OpenConnection())
            using (var transaction = connection.BeginTransaction())
            using (var command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = @"
DELETE FROM local_template_rules
WHERE remote_template_id = @remoteTemplateId
   OR template_name = @templateName
   OR (@sourceRemoteTemplateId <> '' AND remote_template_id = @sourceRemoteTemplateId);";
                command.Parameters.AddWithValue("@exactFingerprint", fingerprint.ExactFingerprint);
                command.Parameters.AddWithValue("@templateName", templateName.Trim());
                command.Parameters.AddWithValue("@remoteTemplateId", templateId.Trim());
                command.Parameters.AddWithValue("@sourceRemoteTemplateId", sourceRemoteTemplateId ?? string.Empty);
                command.ExecuteNonQuery();
                command.Parameters.Clear();

                command.CommandText = @"
INSERT OR REPLACE INTO local_template_rules (
    id, template_name, exact_fingerprint, fuzzy_fingerprint, fingerprint_json, directory_metadata_json,
    rule_json, generation_config_json, status, local_sync_status, sync_error, conflict_remote_json,
    remote_template_id, remote_version, remote_updated_at, deleted_at, updated_at
) VALUES (
    @id, @templateName, @exactFingerprint, @fuzzyFingerprint, @fingerprintJson, @directoryMetadataJson,
    @ruleJson, @generationConfigJson, @status, @localSyncStatus, @syncError, @conflictRemoteJson,
    @remoteTemplateId, @remoteVersion, @remoteUpdatedAt, @deletedAt, @updatedAt
);";
                command.Parameters.AddWithValue("@id", templateId.Trim());
                command.Parameters.AddWithValue("@templateName", templateName.Trim());
                command.Parameters.AddWithValue("@exactFingerprint", fingerprint.ExactFingerprint);
                command.Parameters.AddWithValue("@fuzzyFingerprint", fingerprint.FuzzyFingerprint ?? string.Empty);
                command.Parameters.AddWithValue("@fingerprintJson", JsonConvert.SerializeObject(fingerprint));
                command.Parameters.AddWithValue("@directoryMetadataJson", JsonConvert.SerializeObject(normalizedDirectoryMetadata));
                command.Parameters.AddWithValue("@ruleJson", JsonConvert.SerializeObject(safeRules));
                command.Parameters.AddWithValue("@generationConfigJson", generationConfiguration == null
                    ? string.Empty
                    : JsonConvert.SerializeObject(generationConfiguration));
                command.Parameters.AddWithValue("@status", (int)status);
                command.Parameters.AddWithValue("@localSyncStatus", (int)TemplateSyncStatus.Synced);
                command.Parameters.AddWithValue("@syncError", string.Empty);
                command.Parameters.AddWithValue("@conflictRemoteJson", string.Empty);
                command.Parameters.AddWithValue("@remoteTemplateId", templateId.Trim());
                command.Parameters.AddWithValue("@remoteVersion", Math.Max(0, version));
                command.Parameters.AddWithValue("@remoteUpdatedAt", remoteTime);
                command.Parameters.AddWithValue("@deletedAt", deletedTime);
                command.Parameters.AddWithValue("@updatedAt", now);
                command.ExecuteNonQuery();
                transaction.Commit();
            }
        }

        private void SaveRemoteConflictSnapshot(string exactFingerprint, TemplateConflictRemoteSnapshot snapshot)
        {
            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
UPDATE local_template_rules
SET local_sync_status = @localSyncStatus,
    sync_error = @syncError,
    conflict_remote_json = @conflictRemoteJson,
    updated_at = @updatedAt
WHERE exact_fingerprint = @exactFingerprint;";
                command.Parameters.AddWithValue("@localSyncStatus", (int)TemplateSyncStatus.Conflict);
                command.Parameters.AddWithValue("@syncError", "本地模板存在待上传修改，远端更新未覆盖本地内容。请处理冲突后重新同步。");
                command.Parameters.AddWithValue("@conflictRemoteJson", JsonConvert.SerializeObject(snapshot));
                command.Parameters.AddWithValue("@updatedAt", DateTime.UtcNow.ToString("o"));
                command.Parameters.AddWithValue("@exactFingerprint", exactFingerprint);
                command.ExecuteNonQuery();
            }
        }

        public string GetPreference(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return string.Empty;
            }

            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = "SELECT value FROM user_preferences WHERE key = @key LIMIT 1;";
                command.Parameters.AddWithValue("@key", key);
                return Convert.ToString(command.ExecuteScalar()) ?? string.Empty;
            }
        }

        public void SetPreference(string key, string value)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return;
            }

            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
INSERT OR REPLACE INTO user_preferences (key, value, updated_at)
VALUES (@key, @value, @updatedAt);";
                command.Parameters.AddWithValue("@key", key);
                command.Parameters.AddWithValue("@value", value ?? string.Empty);
                command.Parameters.AddWithValue("@updatedAt", DateTime.UtcNow.ToString("o"));
                command.ExecuteNonQuery();
            }
        }

        private static bool IsOlderRemoteVersion(CachedTemplateRule existing, int version, DateTime? remoteUpdatedAt)
        {
            if (version > 0 && existing.RemoteVersion > version)
            {
                return true;
            }

            if (version > 0 && existing.RemoteVersion == version && remoteUpdatedAt.HasValue && existing.RemoteUpdatedAt.HasValue)
            {
                return existing.RemoteUpdatedAt.Value >= remoteUpdatedAt.Value;
            }

            return version <= 0 && remoteUpdatedAt.HasValue && existing.RemoteUpdatedAt.HasValue &&
                existing.RemoteUpdatedAt.Value >= remoteUpdatedAt.Value;
        }
    }

    public enum TemplateRemoteApplyResult
    {
        Applied,
        Ignored,
        Conflict
    }
}
