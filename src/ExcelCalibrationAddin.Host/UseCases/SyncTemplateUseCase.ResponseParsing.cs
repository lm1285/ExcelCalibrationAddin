using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public sealed partial class SyncTemplateUseCase
    {
        private SyncTemplateResult ParseMatchResult(string raw)
        {
            var root = JObject.Parse(raw);
            var found = root.Value<bool?>("found") ?? false;
            var first = root["templates"]?.First;

            return new SyncTemplateResult
            {
                Found = found,
                TemplateId = first?["id"]?.Value<string>() ?? string.Empty,
                TemplateName = first?["name"]?.Value<string>() ?? string.Empty,
                Version = first?["version"]?.Value<int?>() ?? first?["remote_version"]?.Value<int?>() ?? 0,
                MatchScore = first?["matchScore"]?.Value<double?>() ?? 0,
                MatchReason = first?["matchReason"]?.Value<string>() ?? string.Empty,
                Rules = ParseRules(first?["rules_json"] ?? first?["rules"]),
                GenerationConfiguration = ParseGenerationConfiguration(first?["generation_config_json"] ?? first?["generationConfiguration"]),
                Status = ParseStatus(first?["status"]),
                RawResponse = raw
            };
        }

        private static IReadOnlyList<MeasurementRule> ParseRules(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return new List<MeasurementRule>();
            }

            try
            {
                if (token.Type == JTokenType.String)
                {
                    var raw = token.Value<string>();
                    return string.IsNullOrWhiteSpace(raw)
                        ? new List<MeasurementRule>()
                        : JsonConvert.DeserializeObject<List<MeasurementRule>>(raw) ?? new List<MeasurementRule>();
                }

                return token.ToObject<List<MeasurementRule>>() ?? new List<MeasurementRule>();
            }
            catch
            {
                return new List<MeasurementRule>();
            }
        }

        private static GenerationConfiguration ParseGenerationConfiguration(JToken token)
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
                        : JsonConvert.DeserializeObject<GenerationConfiguration>(raw);
                }

                return token.ToObject<GenerationConfiguration>();
            }
            catch
            {
                return null;
            }
        }

        private static TemplateLifecycleStatus? ParseStatus(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            var text = token.ToString().Trim();
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

            int value;
            return int.TryParse(text, out value) && Enum.IsDefined(typeof(TemplateLifecycleStatus), value)
                ? (TemplateLifecycleStatus?)value
                : null;
        }

    }
}
