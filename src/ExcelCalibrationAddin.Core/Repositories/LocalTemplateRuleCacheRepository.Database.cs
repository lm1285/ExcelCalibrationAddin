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
        private SQLiteConnection OpenConnection()
        {
            var connection = new SQLiteConnection(_connectionString);
            connection.Open();
            using (var command = connection.CreateCommand())
            {
                command.CommandText = "PRAGMA foreign_keys = ON;";
                command.ExecuteNonQuery();
            }
            return connection;
        }

        private static void EnsureColumn(SQLiteConnection connection, string tableName, string columnName, string columnType)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText = $"PRAGMA table_info({tableName});";
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (string.Equals(reader["name"].ToString(), columnName, StringComparison.OrdinalIgnoreCase))
                        {
                            return;
                        }
                    }
                }
            }

            using (var command = connection.CreateCommand())
            {
                command.CommandText = $"ALTER TABLE {tableName} ADD COLUMN {columnName} {columnType};";
                command.ExecuteNonQuery();
            }
        }

        private static TemplateFingerprint DeserializeFingerprint(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            try
            {
                return JsonConvert.DeserializeObject<TemplateFingerprint>(raw);
            }
            catch
            {
                return null;
            }
        }

        private static TemplateDirectoryMetadata DeserializeDirectoryMetadata(string raw, string fallbackTemplateName)
        {
            try
            {
                var metadata = string.IsNullOrWhiteSpace(raw)
                    ? null
                    : JsonConvert.DeserializeObject<TemplateDirectoryMetadata>(raw);
                return NormalizeDirectoryMetadata(metadata, fallbackTemplateName);
            }
            catch
            {
                return NormalizeDirectoryMetadata(null, fallbackTemplateName);
            }
        }

        private static TemplateDirectoryMetadata NormalizeDirectoryMetadata(
            TemplateDirectoryMetadata metadata,
            string fallbackTemplateName)
        {
            metadata = metadata ?? new TemplateDirectoryMetadata();
            return new TemplateDirectoryMetadata
            {
                MeasurementDomain = string.IsNullOrWhiteSpace(metadata.MeasurementDomain)
                    ? "未分类"
                    : metadata.MeasurementDomain.Trim(),
                TemplateName = string.IsNullOrWhiteSpace(metadata.TemplateName)
                    ? (fallbackTemplateName ?? string.Empty).Trim()
                    : metadata.TemplateName.Trim(),
                VariantName = string.IsNullOrWhiteSpace(metadata.VariantName)
                    ? "默认方案"
                    : metadata.VariantName.Trim(),
                TemplateCode = (metadata.TemplateCode ?? string.Empty).Trim()
            };
        }

        public bool UpdateTemplateSyncStatus(string exactFingerprint, TemplateSyncStatus status, string syncError = null)
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
SET local_sync_status = @localSyncStatus,
    sync_error = @syncError,
    updated_at = @updatedAt
WHERE exact_fingerprint = @exactFingerprint;";
                command.Parameters.AddWithValue("@localSyncStatus", (int)status);
                command.Parameters.AddWithValue("@syncError", syncError ?? string.Empty);
                command.Parameters.AddWithValue("@updatedAt", DateTime.UtcNow.ToString("o"));
                command.Parameters.AddWithValue("@exactFingerprint", exactFingerprint);
                return command.ExecuteNonQuery() > 0;
            }
        }

    }
}
