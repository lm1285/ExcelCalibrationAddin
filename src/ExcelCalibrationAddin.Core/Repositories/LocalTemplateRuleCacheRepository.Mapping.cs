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
        private static GenerationConfiguration DeserializeGenerationConfiguration(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            try
            {
                return JsonConvert.DeserializeObject<GenerationConfiguration>(raw);
            }
            catch
            {
                return null;
            }
        }

        private static CachedTemplateRule ReadCachedTemplateRule(SQLiteDataReader reader)
        {
            return new CachedTemplateRule
            {
                TemplateName = reader["template_name"].ToString(),
                DirectoryMetadata = DeserializeDirectoryMetadata(
                    reader["directory_metadata_json"].ToString(),
                    reader["template_name"].ToString()),
                ExactFingerprint = reader["exact_fingerprint"].ToString(),
                FuzzyFingerprint = reader["fuzzy_fingerprint"].ToString(),
                Fingerprint = DeserializeFingerprint(reader["fingerprint_json"].ToString()),
                Rules = JsonConvert.DeserializeObject<List<MeasurementRule>>(reader["rule_json"].ToString() ?? "[]") ?? new List<MeasurementRule>(),
                GenerationConfiguration = DeserializeGenerationConfiguration(reader["generation_config_json"].ToString()),
                Status = ReadTemplateStatus(reader["status"]),
                LocalSyncStatus = ReadTemplateSyncStatus(reader["local_sync_status"]),
                SyncError = reader["sync_error"].ToString(),
                HasRemoteConflict = !string.IsNullOrWhiteSpace(reader["conflict_remote_json"].ToString()),
                RemoteTemplateId = reader["remote_template_id"].ToString(),
                RemoteVersion = ReadInt(reader["remote_version"]),
                RemoteUpdatedAt = ReadNullableDate(reader["remote_updated_at"]),
                DeletedAt = ReadNullableDate(reader["deleted_at"]),
                UpdatedAt = DateTime.Parse(reader["updated_at"].ToString() ?? DateTime.UtcNow.ToString("o"))
            };
        }

        private static int ReadInt(object value)
        {
            try
            {
                return Convert.ToInt32(value);
            }
            catch
            {
                return 0;
            }
        }

        private static DateTime? ReadNullableDate(object value)
        {
            DateTime parsed;
            return DateTime.TryParse(value?.ToString(), out parsed) ? parsed : (DateTime?)null;
        }

        private static TemplateLifecycleStatus ReadTemplateStatus(object raw)
        {
            try
            {
                return Enum.IsDefined(typeof(TemplateLifecycleStatus), Convert.ToInt32(raw))
                    ? (TemplateLifecycleStatus)Convert.ToInt32(raw)
                    : TemplateLifecycleStatus.Enabled;
            }
            catch
            {
                return TemplateLifecycleStatus.Enabled;
            }
        }

        private static double ScoreFingerprint(TemplateFingerprint current, TemplateFingerprint stored, string storedFuzzyFingerprint)
        {
            if (current == null)
            {
                return 0;
            }

            if (!string.IsNullOrWhiteSpace(current.FuzzyFingerprint) &&
                string.Equals(current.FuzzyFingerprint, storedFuzzyFingerprint, StringComparison.OrdinalIgnoreCase))
            {
                return 75;
            }

            if (stored == null)
            {
                return 0;
            }

            var titleScore = Similarity(current.Title, stored.Title);
            var sheetScore = OverlapRatio(current.SheetNames, stored.SheetNames);
            var headerScore = OverlapRatio(current.HeaderTexts, stored.HeaderTexts);
            var hasSheets = (current.SheetNames?.Count ?? 0) > 0 && (stored.SheetNames?.Count ?? 0) > 0;
            return hasSheets
                ? Math.Round(titleScore * 30 + sheetScore * 25 + headerScore * 45, 0)
                : Math.Round(titleScore * 35 + headerScore * 65, 0);
        }

        private static bool IsLegacyCompatibleFingerprint(
            TemplateFingerprint current,
            TemplateFingerprint stored,
            string storedFuzzyFingerprint)
        {
            if (current == null || stored == null ||
                !string.IsNullOrWhiteSpace(stored.StructureSignature) ||
                string.IsNullOrWhiteSpace(current.FuzzyFingerprint) ||
                !string.Equals(current.FuzzyFingerprint, storedFuzzyFingerprint, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return string.Equals(Normalize(current.Title), Normalize(stored.Title), StringComparison.Ordinal) &&
                HasSameTextSet(current.SheetNames, stored.SheetNames) &&
                HasSameTextSet(current.HeaderTexts, stored.HeaderTexts);
        }

        private static bool IsStructureCompatibleFingerprint(
            TemplateFingerprint current,
            TemplateFingerprint stored)
        {
            return current != null &&
                stored != null &&
                !string.IsNullOrWhiteSpace(current.StructureSignature) &&
                !string.IsNullOrWhiteSpace(stored.StructureSignature) &&
                string.Equals(
                    current.StructureSignature,
                    stored.StructureSignature,
                    StringComparison.Ordinal) &&
                HasSameTextSet(current.SheetNames, stored.SheetNames);
        }

        private static bool HasSameTextSet(IReadOnlyList<string> left, IReadOnlyList<string> right)
        {
            var leftSet = new HashSet<string>((left ?? new List<string>()).Select(Normalize), StringComparer.OrdinalIgnoreCase);
            var rightSet = new HashSet<string>((right ?? new List<string>()).Select(Normalize), StringComparer.OrdinalIgnoreCase);
            leftSet.RemoveWhere(string.IsNullOrWhiteSpace);
            rightSet.RemoveWhere(string.IsNullOrWhiteSpace);
            return leftSet.Count > 0 && leftSet.SetEquals(rightSet);
        }

        private static double OverlapRatio(IReadOnlyList<string> left, IReadOnlyList<string> right)
        {
            var leftSet = new HashSet<string>((left ?? new List<string>()).Select(Normalize), StringComparer.OrdinalIgnoreCase);
            var rightSet = new HashSet<string>((right ?? new List<string>()).Select(Normalize), StringComparer.OrdinalIgnoreCase);
            leftSet.RemoveWhere(string.IsNullOrWhiteSpace);
            rightSet.RemoveWhere(string.IsNullOrWhiteSpace);
            if (leftSet.Count == 0 || rightSet.Count == 0)
            {
                return 0;
            }

            var matches = 0;
            foreach (var item in leftSet)
            {
                if (rightSet.Contains(item))
                {
                    matches++;
                }
            }

            return matches / (double)Math.Max(leftSet.Count, rightSet.Count);
        }

        private static double Similarity(string left, string right)
        {
            var normalizedLeft = Normalize(left);
            var normalizedRight = Normalize(right);
            if (string.IsNullOrWhiteSpace(normalizedLeft) || string.IsNullOrWhiteSpace(normalizedRight))
            {
                return 0;
            }

            if (string.Equals(normalizedLeft, normalizedRight, StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            if (normalizedLeft.Contains(normalizedRight) || normalizedRight.Contains(normalizedLeft))
            {
                return 0.9;
            }

            return 0;
        }

        private static string Normalize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var chars = new List<char>();
            foreach (var ch in value.Trim())
            {
                if (!char.IsWhiteSpace(ch) && ch != '_' && ch != '-' && ch != '/' && ch != ':' && ch != ',' && ch != '.' && ch != '(' && ch != ')' && ch != '[' && ch != ']')
                {
                    chars.Add(char.ToLowerInvariant(ch));
                }
            }

            return new string(chars.ToArray());
        }

        private static TemplateSyncStatus ReadTemplateSyncStatus(object raw)
        {
            try
            {
                return Enum.IsDefined(typeof(TemplateSyncStatus), Convert.ToInt32(raw))
                    ? (TemplateSyncStatus)Convert.ToInt32(raw)
                    : TemplateSyncStatus.Synced;
            }
            catch
            {
                return TemplateSyncStatus.Synced;
            }
        }

    }
}
