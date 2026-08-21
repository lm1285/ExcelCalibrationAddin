using System;
using System.Collections.Generic;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelCalibrationAddin.Host.Templates
{
    public static class RemoteTemplateSaveResponseParser
    {
        public static RemoteTemplateSnapshot Parse(
            string raw,
            string fallbackTemplateName,
            TemplateFingerprint fallbackFingerprint,
            IReadOnlyList<MeasurementRule> fallbackRules,
            GenerationConfiguration fallbackGenerationConfiguration)
        {
            var token = ResolvePayloadToken(raw);
            var snapshot = new RemoteTemplateSnapshot
            {
                TemplateId = ReadString(token, "template_id", "id", "remote_template_id") ?? string.Empty,
                TemplateName = ReadString(token, "template_name", "name") ?? string.Empty,
                Version = ReadInt(token, "version", "remote_version"),
                RemoteUpdatedAt = ReadDate(token, "updated_at", "remote_updated_at"),
                DeletedAt = ReadDate(token, "deleted_at"),
                Status = ReadStatus(token?["status"]) ?? TemplateLifecycleStatus.Enabled,
                Fingerprint = ReadFingerprint(token),
                Rules = ReadRules(token),
                GenerationConfiguration = ReadGenerationConfiguration(token),
                DirectoryMetadata = ReadDirectoryMetadata(token)
            };

            Validate(snapshot);
            return snapshot;
        }

        private static void Validate(RemoteTemplateSnapshot snapshot)
        {
            if (snapshot == null ||
                string.IsNullOrWhiteSpace(snapshot.TemplateId) ||
                string.IsNullOrWhiteSpace(snapshot.TemplateName) ||
                string.IsNullOrWhiteSpace(snapshot.Fingerprint?.ExactFingerprint) ||
                snapshot.Rules == null ||
                snapshot.Rules.Count == 0)
            {
                throw new InvalidOperationException("远端模板服务未回传完整模板数据，已转为本地待上传状态。");
            }
        }

        private static JToken ResolvePayloadToken(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return new JObject();
            }

            try
            {
                var root = JToken.Parse(raw);
                return root["data"] ?? root["template"] ?? root;
            }
            catch
            {
                return new JObject();
            }
        }

        private static string ReadString(JToken token, params string[] names)
        {
            foreach (var name in names)
            {
                var value = token?[name]?.Value<string>();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return null;
        }

        private static int ReadInt(JToken token, params string[] names)
        {
            foreach (var name in names)
            {
                var value = token?[name]?.Value<int?>();
                if (value.HasValue)
                {
                    return value.Value;
                }
            }

            return 0;
        }

        private static DateTime? ReadDate(JToken token, params string[] names)
        {
            foreach (var name in names)
            {
                var text = token?[name]?.Value<string>();
                if (DateTime.TryParse(text, out var value))
                {
                    return value.ToUniversalTime();
                }
            }

            return null;
        }

        private static TemplateLifecycleStatus? ReadStatus(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            var text = token.ToString().Trim();
            if (int.TryParse(text, out var numeric) && Enum.IsDefined(typeof(TemplateLifecycleStatus), numeric))
            {
                return (TemplateLifecycleStatus)numeric;
            }

            if (string.Equals(text, "enabled", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "启用", StringComparison.OrdinalIgnoreCase))
            {
                return TemplateLifecycleStatus.Enabled;
            }

            if (string.Equals(text, "disabled", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "停用", StringComparison.OrdinalIgnoreCase))
            {
                return TemplateLifecycleStatus.Disabled;
            }

            if (string.Equals(text, "obsolete", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "deprecated", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "废止", StringComparison.OrdinalIgnoreCase))
            {
                return TemplateLifecycleStatus.Obsolete;
            }

            return null;
        }

        private static TemplateFingerprint ReadFingerprint(JToken token)
        {
            return ReadJsonValue<TemplateFingerprint>(token?["fingerprint"] ?? token?["fingerprint_hash"]);
        }

        private static IReadOnlyList<MeasurementRule> ReadRules(JToken token)
        {
            return ReadJsonValue<List<MeasurementRule>>(token?["rules"] ?? token?["rules_json"]);
        }

        private static GenerationConfiguration ReadGenerationConfiguration(JToken token)
        {
            return ReadJsonValue<GenerationConfiguration>(token?["generationConfiguration"] ?? token?["generation_config_json"]);
        }

        private static TemplateDirectoryMetadata ReadDirectoryMetadata(JToken token)
        {
            return ReadJsonValue<TemplateDirectoryMetadata>(token?["directoryMetadata"] ?? token?["directory_metadata"]);
        }

        private static T ReadJsonValue<T>(JToken token) where T : class
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            try
            {
                if (token.Type == JTokenType.String)
                {
                    var raw = token.Value<string>();
                    return string.IsNullOrWhiteSpace(raw)
                        ? null
                        : JsonConvert.DeserializeObject<T>(raw);
                }

                return token.ToObject<T>();
            }
            catch
            {
                return null;
            }
        }
    }

    public sealed class RemoteTemplateSnapshot
    {
        public string TemplateId { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public int Version { get; set; } = 1;
        public DateTime? RemoteUpdatedAt { get; set; }
        public DateTime? DeletedAt { get; set; }
        public TemplateLifecycleStatus Status { get; set; } = TemplateLifecycleStatus.Enabled;
        public TemplateFingerprint Fingerprint { get; set; }
        public IReadOnlyList<MeasurementRule> Rules { get; set; } = new List<MeasurementRule>();
        public GenerationConfiguration GenerationConfiguration { get; set; }
        public TemplateDirectoryMetadata DirectoryMetadata { get; set; }
    }
}
