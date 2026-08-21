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
        public bool ResolveTemplateConflict(string exactFingerprint, TemplateConflictResolutionAction action, string saveAsTemplateName = null)
        {
            var local = FindByExactFingerprint(exactFingerprint);
            if (local == null || local.LocalSyncStatus != TemplateSyncStatus.Conflict)
            {
                return false;
            }

            var remote = LoadConflictRemoteSnapshot(exactFingerprint);
            if (remote == null && action != TemplateConflictResolutionAction.KeepLocal)
            {
                throw new InvalidOperationException("冲突模板缺少远端副本，请重新同步后再处理。");
            }

            switch (action)
            {
                case TemplateConflictResolutionAction.KeepLocal:
                    return KeepLocalConflictVersion(local.ExactFingerprint);
                case TemplateConflictResolutionAction.UseRemote:
                    ApplyRemoteConflictSnapshot(remote);
                    return true;
                case TemplateConflictResolutionAction.SaveAs:
                    if (string.IsNullOrWhiteSpace(saveAsTemplateName))
                    {
                        throw new InvalidOperationException("另存冲突模板时必须提供新模板名称。");
                    }

                    ApplyRemoteConflictSnapshot(remote);
                    SaveConflictLocalCopy(local, saveAsTemplateName);
                    return true;
                default:
                    throw new ArgumentOutOfRangeException(nameof(action), action, "未知的冲突处理动作。");
            }
        }

        internal TemplateConflictRemoteSnapshot LoadConflictRemoteSnapshot(string exactFingerprint)
        {
            if (string.IsNullOrWhiteSpace(exactFingerprint))
            {
                return null;
            }

            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT conflict_remote_json
FROM local_template_rules
WHERE exact_fingerprint = @exactFingerprint
ORDER BY updated_at DESC
LIMIT 1;";
                command.Parameters.AddWithValue("@exactFingerprint", exactFingerprint);
                var raw = Convert.ToString(command.ExecuteScalar());
                if (string.IsNullOrWhiteSpace(raw))
                {
                    return null;
                }

                return JsonConvert.DeserializeObject<TemplateConflictRemoteSnapshot>(raw);
            }
        }

        private bool KeepLocalConflictVersion(string exactFingerprint)
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
                command.Parameters.AddWithValue("@localSyncStatus", (int)TemplateSyncStatus.PendingUpload);
                command.Parameters.AddWithValue("@syncError", "用户选择保留本地版本，等待下次上传。");
                command.Parameters.AddWithValue("@conflictRemoteJson", string.Empty);
                command.Parameters.AddWithValue("@updatedAt", DateTime.UtcNow.ToString("o"));
                command.Parameters.AddWithValue("@exactFingerprint", exactFingerprint);
                return command.ExecuteNonQuery() > 0;
            }
        }

        private void SaveConflictLocalCopy(CachedTemplateRule local, string saveAsTemplateName)
        {
            SaveTemplate(
                saveAsTemplateName.Trim(),
                local.Fingerprint ?? new TemplateFingerprint
                {
                    ExactFingerprint = local.ExactFingerprint,
                    FuzzyFingerprint = local.FuzzyFingerprint
                },
                local.Rules,
                local.GenerationConfiguration,
                createNew: true,
                localSyncStatus: TemplateSyncStatus.PendingUpload,
                syncError: "用户从冲突模板另存，等待上传为新模板。");
        }

        private void ApplyRemoteConflictSnapshot(TemplateConflictRemoteSnapshot remote)
        {
            if (remote?.Fingerprint == null || string.IsNullOrWhiteSpace(remote.Fingerprint.ExactFingerprint))
            {
                throw new InvalidOperationException("冲突远端副本缺少有效指纹。");
            }

            SaveRemoteTemplate(
                remote.TemplateId,
                remote.TemplateName,
                remote.Version,
                remote.RemoteUpdatedAt,
                remote.DeletedAt,
                remote.Status,
                remote.Fingerprint,
                remote.Rules,
                remote.GenerationConfiguration);
        }

        private static TemplateConflictRemoteSnapshot BuildConflictSnapshot(
            string templateId,
            string templateName,
            int version,
            DateTime? remoteUpdatedAt,
            DateTime? deletedAt,
            TemplateLifecycleStatus status,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration)
        {
            return new TemplateConflictRemoteSnapshot
            {
                TemplateId = templateId ?? string.Empty,
                TemplateName = templateName ?? string.Empty,
                Version = version,
                RemoteUpdatedAt = remoteUpdatedAt,
                DeletedAt = deletedAt,
                Status = status,
                Fingerprint = fingerprint,
                Rules = (rules ?? new List<MeasurementRule>()).Where(item => item != null).Select(CloneTemplateRule).ToList(),
                GenerationConfiguration = generationConfiguration
            };
        }
    }

    public enum TemplateConflictResolutionAction
    {
        KeepLocal,
        UseRemote,
        SaveAs
    }

    internal sealed class TemplateConflictRemoteSnapshot
    {
        public string TemplateId { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public int Version { get; set; }
        public DateTime? RemoteUpdatedAt { get; set; }
        public DateTime? DeletedAt { get; set; }
        public TemplateLifecycleStatus Status { get; set; } = TemplateLifecycleStatus.Enabled;
        public TemplateFingerprint Fingerprint { get; set; }
        public List<MeasurementRule> Rules { get; set; } = new List<MeasurementRule>();
        public GenerationConfiguration GenerationConfiguration { get; set; }
    }
}
