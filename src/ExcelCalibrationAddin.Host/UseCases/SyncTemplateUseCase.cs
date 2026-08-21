using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public sealed partial class SyncTemplateUseCase
    {
        private readonly TemplateSyncClient _syncClient;
        private readonly LocalTemplateRuleCacheRepository _cacheRepository;

        public SyncTemplateUseCase(TemplateSyncClient syncClient, LocalTemplateRuleCacheRepository cacheRepository)
        {
            _syncClient = syncClient;
            _cacheRepository = cacheRepository;
        }

        public async Task<SyncTemplateResult> MatchRemoteAsync(TemplateFingerprint fingerprint)
        {
            try
            {
                return ParseMatchResult(await _syncClient.MatchAsync(fingerprint));
            }
            catch (Exception ex)
            {
                return new SyncTemplateResult
                {
                    Found = false,
                    ErrorMessage = ex.Message
                };
            }
        }

    }

    public sealed class SyncTemplateResult
    {
        public bool Found { get; set; }
        public string TemplateId { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public int Version { get; set; }
        public double MatchScore { get; set; }
        public string MatchReason { get; set; } = string.Empty;
        public string RawResponse { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
        public IReadOnlyList<MeasurementRule> Rules { get; set; } = new List<MeasurementRule>();
        public GenerationConfiguration GenerationConfiguration { get; set; }
        public TemplateLifecycleStatus? Status { get; set; }
    }
}
