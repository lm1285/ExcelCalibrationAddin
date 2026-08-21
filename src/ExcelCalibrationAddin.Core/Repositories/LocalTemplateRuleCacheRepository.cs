using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Repositories
{
    public sealed partial class LocalTemplateRuleCacheRepository
    {
        private readonly string _connectionString;

        public LocalTemplateRuleCacheRepository(string sqliteFilePath)
        {
            var fullPath = Path.GetFullPath(sqliteFilePath);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? ".");
            _connectionString = new SQLiteConnectionStringBuilder
            {
                DataSource = fullPath,
                Version = 3
            }.ConnectionString;
        }

        public void Initialize()
        {
            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = Services.LocalRuleCacheSchema.CreateRulesTable;
                command.ExecuteNonQuery();
                EnsureColumn(connection, "local_template_rules", "fingerprint_json", "TEXT");
                EnsureColumn(connection, "local_template_rules", "directory_metadata_json", "TEXT");
                EnsureColumn(connection, "local_template_rules", "generation_config_json", "TEXT");
                EnsureColumn(connection, "local_template_rules", "status", "INTEGER NOT NULL DEFAULT 0");
                EnsureColumn(connection, "local_template_rules", "local_sync_status", "INTEGER NOT NULL DEFAULT 0");
                EnsureColumn(connection, "local_template_rules", "sync_error", "TEXT");
                EnsureColumn(connection, "local_template_rules", "conflict_remote_json", "TEXT");
                EnsureColumn(connection, "local_template_rules", "remote_updated_at", "TEXT");
                EnsureColumn(connection, "local_template_rules", "deleted_at", "TEXT");
                command.CommandText = Services.LocalRuleCacheSchema.CreatePreferencesTable;
                command.ExecuteNonQuery();
                command.CommandText = Services.LocalRuleCacheSchema.CreateSampleDataTables;
                command.ExecuteNonQuery();
                command.CommandText = "PRAGMA foreign_keys = ON;";
                command.ExecuteNonQuery();
            }
        }

        public CachedTemplateRule FindByExactFingerprint(string exactFingerprint)
        {
            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT template_name, exact_fingerprint, fuzzy_fingerprint, fingerprint_json, directory_metadata_json, rule_json, generation_config_json, status, local_sync_status, sync_error, conflict_remote_json, remote_template_id, remote_version, remote_updated_at, deleted_at, updated_at
FROM local_template_rules
WHERE exact_fingerprint = @exactFingerprint
ORDER BY updated_at DESC
LIMIT 1;";
                command.Parameters.AddWithValue("@exactFingerprint", exactFingerprint);

                using (var reader = command.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        return null;
                    }

                    return ReadCachedTemplateRule(reader);
                }
            }
        }

        public CachedTemplateRule FindByRemoteTemplateId(string remoteTemplateId)
        {
            if (string.IsNullOrWhiteSpace(remoteTemplateId))
            {
                return null;
            }

            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT template_name, exact_fingerprint, fuzzy_fingerprint, fingerprint_json, directory_metadata_json, rule_json, generation_config_json, status, local_sync_status, sync_error, conflict_remote_json, remote_template_id, remote_version, remote_updated_at, deleted_at, updated_at
FROM local_template_rules
WHERE remote_template_id = @remoteTemplateId
ORDER BY updated_at DESC
LIMIT 1;";
                command.Parameters.AddWithValue("@remoteTemplateId", remoteTemplateId);
                using (var reader = command.ExecuteReader())
                {
                    return reader.Read() ? ReadCachedTemplateRule(reader) : null;
                }
            }
        }

        public CachedTemplateRule FindByExactOrLegacyCompatibleFingerprint(TemplateFingerprint fingerprint)
        {
            var exact = FindByExactFingerprint(fingerprint?.ExactFingerprint ?? string.Empty);
            if (exact != null)
            {
                exact.MatchScore = 100;
                exact.MatchReason = "exact fingerprint";
                return exact;
            }

            foreach (var candidate in ListSavedTemplates())
            {
                if (IsStructureCompatibleFingerprint(fingerprint, candidate?.Fingerprint))
                {
                    candidate.MatchScore = 100;
                    candidate.MatchReason = "structure signature";
                    return candidate;
                }

                if (!IsLegacyCompatibleFingerprint(fingerprint, candidate?.Fingerprint, candidate?.FuzzyFingerprint))
                {
                    continue;
                }

                candidate.MatchScore = 100;
                candidate.MatchReason = "legacy compatible fingerprint";
                return candidate;
            }

            return null;
        }

        public CachedTemplateRule FindBestMatch(TemplateFingerprint fingerprint, double minimumScore = 60)
        {
            var exact = FindByExactFingerprint(fingerprint?.ExactFingerprint ?? string.Empty);
            if (exact != null)
            {
                exact.MatchScore = 100;
                exact.MatchReason = "exact fingerprint";
                return exact;
            }

            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT template_name, exact_fingerprint, fuzzy_fingerprint, fingerprint_json, directory_metadata_json, rule_json, generation_config_json, status, local_sync_status, sync_error, conflict_remote_json, remote_template_id, remote_version, remote_updated_at, deleted_at, updated_at
FROM local_template_rules
ORDER BY updated_at DESC;";

                CachedTemplateRule best = null;
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var candidate = ReadCachedTemplateRule(reader);

                        var score = ScoreFingerprint(fingerprint, candidate.Fingerprint, candidate.FuzzyFingerprint);
                        if (score < minimumScore || (best != null && score <= best.MatchScore))
                        {
                            continue;
                        }

                        candidate.MatchScore = score;
                        candidate.MatchReason = "local fingerprint similarity";
                        best = candidate;
                    }
                }

                return best;
            }
        }

        public IReadOnlyList<CachedTemplateRule> ListSavedTemplates()
        {
            var templates = new List<CachedTemplateRule>();
            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT template_name, exact_fingerprint, fuzzy_fingerprint, fingerprint_json, directory_metadata_json, rule_json, generation_config_json, status, local_sync_status, sync_error, conflict_remote_json, remote_template_id, remote_version, remote_updated_at, deleted_at, updated_at
FROM local_template_rules
ORDER BY updated_at DESC;";

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        templates.Add(ReadCachedTemplateRule(reader));
                    }
                }
            }

            return templates;
        }

        public IReadOnlyList<CachedTemplateRule> ListTemplatesBySyncStatus(params TemplateSyncStatus[] statuses)
        {
            var statusSet = new HashSet<TemplateSyncStatus>(statuses ?? Array.Empty<TemplateSyncStatus>());
            return ListSavedTemplates()
                .Where(template => statusSet.Contains(template.LocalSyncStatus))
                .ToList();
        }

        public void SaveTemplate(
            string templateName,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration = null,
            bool createNew = false,
            TemplateSyncStatus localSyncStatus = TemplateSyncStatus.Synced,
            string syncError = null,
            TemplateDirectoryMetadata directoryMetadata = null,
            string targetRemoteTemplateId = null)
        {
            if (string.IsNullOrWhiteSpace(templateName))
            {
                throw new ArgumentException("模板名称不能为空。", nameof(templateName));
            }

            if (string.IsNullOrWhiteSpace(fingerprint?.ExactFingerprint))
            {
                throw new InvalidOperationException("当前识别结果缺少有效模板指纹，无法保存。");
            }

            var safeRules = (rules ?? new List<MeasurementRule>())
                .Where(rule => rule != null)
                .Select(CloneTemplateRule)
                .ToList();
            ValidateSavableTemplate(safeRules);

            var now = DateTime.UtcNow.ToString("o");
            var normalizedDirectoryMetadata = NormalizeDirectoryMetadata(directoryMetadata, templateName);
            var previous = string.IsNullOrWhiteSpace(targetRemoteTemplateId)
                ? FindByExactFingerprint(fingerprint.ExactFingerprint)
                : FindByRemoteTemplateId(targetRemoteTemplateId);
            var remoteTemplateId = createNew
                ? $"template:{Guid.NewGuid():N}"
                : string.IsNullOrWhiteSpace(previous?.RemoteTemplateId)
                    ? fingerprint.ExactFingerprint
                    : previous.RemoteTemplateId;
            var remoteVersion = Math.Max(1, (previous?.RemoteVersion ?? 0) + 1);
            using (var connection = OpenConnection())
            using (var transaction = connection.BeginTransaction())
            using (var command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                if (!createNew && !string.IsNullOrWhiteSpace(targetRemoteTemplateId))
                {
                    command.CommandText = "DELETE FROM local_template_rules WHERE remote_template_id = @targetRemoteTemplateId;";
                    command.Parameters.AddWithValue("@targetRemoteTemplateId", targetRemoteTemplateId);
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();
                }

                if (!createNew)
                {
                    command.CommandText = @"
DELETE FROM local_template_rules
WHERE template_name = @templateName
  AND exact_fingerprint <> @exactFingerprint;";
                    command.Parameters.AddWithValue("@templateName", templateName.Trim());
                    command.Parameters.AddWithValue("@exactFingerprint", fingerprint.ExactFingerprint);
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();
                }

                command.CommandText = @"
INSERT OR REPLACE INTO local_template_rules (
    id,
    template_name,
    exact_fingerprint,
    fuzzy_fingerprint,
    fingerprint_json,
    directory_metadata_json,
    rule_json,
    generation_config_json,
    status,
    local_sync_status,
    sync_error,
    conflict_remote_json,
    remote_template_id,
    remote_version,
    remote_updated_at,
    deleted_at,
    updated_at
) VALUES (
    @id,
    @templateName,
    @exactFingerprint,
    @fuzzyFingerprint,
    @fingerprintJson,
    @directoryMetadataJson,
    @ruleJson,
    @generationConfigurationJson,
    @status,
    @localSyncStatus,
    @syncError,
    @conflictRemoteJson,
    @remoteTemplateId,
    @remoteVersion,
    @remoteUpdatedAt,
    @deletedAt,
    @updatedAt
);";
                command.Parameters.AddWithValue("@id", createNew
                    ? $"{fingerprint.ExactFingerprint}:{Guid.NewGuid():N}"
                    : string.IsNullOrWhiteSpace(previous?.RemoteTemplateId)
                        ? fingerprint.ExactFingerprint
                        : previous.RemoteTemplateId);
                command.Parameters.AddWithValue("@templateName", templateName.Trim());
                command.Parameters.AddWithValue("@exactFingerprint", fingerprint.ExactFingerprint);
                command.Parameters.AddWithValue("@fuzzyFingerprint", fingerprint.FuzzyFingerprint ?? string.Empty);
                command.Parameters.AddWithValue("@fingerprintJson", JsonConvert.SerializeObject(fingerprint));
                command.Parameters.AddWithValue("@directoryMetadataJson", JsonConvert.SerializeObject(normalizedDirectoryMetadata));
                command.Parameters.AddWithValue("@ruleJson", JsonConvert.SerializeObject(safeRules));
                command.Parameters.AddWithValue(
                    "@generationConfigurationJson",
                    generationConfiguration == null ? string.Empty : JsonConvert.SerializeObject(generationConfiguration));
                command.Parameters.AddWithValue("@status", (int)TemplateLifecycleStatus.Enabled);
                command.Parameters.AddWithValue("@localSyncStatus", (int)localSyncStatus);
                command.Parameters.AddWithValue("@syncError", syncError ?? string.Empty);
                command.Parameters.AddWithValue("@conflictRemoteJson", string.Empty);
                command.Parameters.AddWithValue("@remoteTemplateId", remoteTemplateId);
                command.Parameters.AddWithValue("@remoteVersion", remoteVersion);
                command.Parameters.AddWithValue("@remoteUpdatedAt", now);
                command.Parameters.AddWithValue("@deletedAt", string.Empty);
                command.Parameters.AddWithValue("@updatedAt", now);
                command.ExecuteNonQuery();
                transaction.Commit();
            }
        }

        public bool UpdateTemplateGenerationConfiguration(string exactFingerprint, GenerationConfiguration generationConfiguration)
        {
            if (string.IsNullOrWhiteSpace(exactFingerprint))
            {
                return false;
            }

            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
UPDATE local_template_rules
SET generation_config_json = @generationConfigurationJson,
    updated_at = @updatedAt
WHERE exact_fingerprint = @exactFingerprint;";
                command.Parameters.AddWithValue(
                    "@generationConfigurationJson",
                    generationConfiguration == null ? string.Empty : JsonConvert.SerializeObject(generationConfiguration));
                command.Parameters.AddWithValue("@updatedAt", DateTime.UtcNow.ToString("o"));
                command.Parameters.AddWithValue("@exactFingerprint", exactFingerprint);
                return command.ExecuteNonQuery() > 0;
            }
        }

        public bool UpdateTemplateStatus(string exactFingerprint, TemplateLifecycleStatus status)
        {
            if (string.IsNullOrWhiteSpace(exactFingerprint))
            {
                return false;
            }

            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
UPDATE local_template_rules
SET status = @status,
    updated_at = @updatedAt
WHERE exact_fingerprint = @exactFingerprint;";
                command.Parameters.AddWithValue("@status", (int)status);
                command.Parameters.AddWithValue("@updatedAt", DateTime.UtcNow.ToString("o"));
                command.Parameters.AddWithValue("@exactFingerprint", exactFingerprint);
                return command.ExecuteNonQuery() > 0;
            }
        }

        public bool DeleteTemplate(string exactFingerprint)
        {
            if (string.IsNullOrWhiteSpace(exactFingerprint))
            {
                return false;
            }

            using (var connection = OpenConnection())
            using (var command = connection.CreateCommand())
            {
                command.CommandText = "DELETE FROM local_template_rules WHERE exact_fingerprint = @exactFingerprint;";
                command.Parameters.AddWithValue("@exactFingerprint", exactFingerprint);
                return command.ExecuteNonQuery() > 0;
            }
        }

    }

    public sealed class CachedTemplateRule
    {
        public string TemplateName { get; set; } = string.Empty;
        public TemplateDirectoryMetadata DirectoryMetadata { get; set; } = new TemplateDirectoryMetadata();
        public string ExactFingerprint { get; set; } = string.Empty;
        public string FuzzyFingerprint { get; set; } = string.Empty;
        public TemplateFingerprint Fingerprint { get; set; }
        public List<MeasurementRule> Rules { get; set; } = new List<MeasurementRule>();
        public GenerationConfiguration GenerationConfiguration { get; set; }
        public string RemoteTemplateId { get; set; } = string.Empty;
        public int RemoteVersion { get; set; }
        public DateTime? RemoteUpdatedAt { get; set; }
        public DateTime? DeletedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
        public TemplateLifecycleStatus Status { get; set; } = TemplateLifecycleStatus.Enabled;
        public TemplateSyncStatus LocalSyncStatus { get; set; } = TemplateSyncStatus.Synced;
        public string SyncError { get; set; } = string.Empty;
        public bool HasRemoteConflict { get; set; }
        public double MatchScore { get; set; }
        public string MatchReason { get; set; } = string.Empty;
    }
}
