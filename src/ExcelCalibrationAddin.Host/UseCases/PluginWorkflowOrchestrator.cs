using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public sealed class PluginWorkflowOrchestrator
    {
        private readonly TemplateRecognitionUseCase _recognitionUseCase;
        private readonly GenerateMeasurementUseCase _generationUseCase;
        private readonly SyncTemplateUseCase _syncTemplateUseCase;
        private readonly LocalTemplateRuleCacheRepository _cacheRepository;

        public PluginWorkflowOrchestrator(
            TemplateRecognitionUseCase recognitionUseCase,
            GenerateMeasurementUseCase generationUseCase,
            SyncTemplateUseCase syncTemplateUseCase,
            LocalTemplateRuleCacheRepository cacheRepository)
        {
            _recognitionUseCase = recognitionUseCase;
            _generationUseCase = generationUseCase;
            _syncTemplateUseCase = syncTemplateUseCase;
            _cacheRepository = cacheRepository;
        }

        public async Task<RecognitionAndSyncResult> RecognizeAndMatchAsync()
        {
            RecognitionProgress.Report(5, "开始读取当前工作表...");
            var recognition = _recognitionUseCase.Execute();
            RecognitionProgress.Report(65, "正在匹配本地模板缓存...");
            var local = _cacheRepository.FindByExactOrLegacyCompatibleFingerprint(recognition.Fingerprint);
            var remote = new SyncTemplateResult { Found = false };
            if (!HasEnabledLocalRules(local))
            {
                RecognitionProgress.Report(80, "正在匹配远端模板库...");
                remote = await _syncTemplateUseCase.MatchRemoteAsync(recognition.Fingerprint);
            }
            RecognitionProgress.Report(100, "识别完成");

            return new RecognitionAndSyncResult
            {
                Recognition = recognition,
                Remote = remote,
                Local = local
            };
        }

        public RecognitionAndSyncResult RecognizeAndMatchLocal()
        {
            RecognitionProgress.Report(5, "开始读取当前工作表...");
            var recognition = _recognitionUseCase.Execute();
            RecognitionProgress.Report(65, "正在匹配本地模板缓存...");
            var local = _cacheRepository.FindByExactOrLegacyCompatibleFingerprint(recognition.Fingerprint);
            RecognitionProgress.Report(100, "识别完成");

            return new RecognitionAndSyncResult
            {
                Recognition = recognition,
                Remote = new SyncTemplateResult { Found = false },
                Local = local
            };
        }

        private static bool HasEnabledLocalRules(CachedTemplateRule local)
        {
            return local != null &&
                local.Status == TemplateLifecycleStatus.Enabled &&
                local.MatchScore >= 100 &&
                local.Rules != null &&
                local.Rules.Count > 0;
        }

        public GenerationWriteResult WriteGeneration(IReadOnlyList<MeasurementRule> rules)
        {
            return _generationUseCase.Write(rules);
        }

        public GenerationWriteResult WriteGeneration(IReadOnlyList<MeasurementRule> rules, GenerationConfiguration generationConfiguration)
        {
            _generationUseCase.SetGenerationConfiguration(generationConfiguration);
            return _generationUseCase.Write(rules);
        }

        public GenerationWriteResult WriteResolvedGeneration(IReadOnlyList<MeasurementRule> rules, GenerationConfiguration generationConfiguration)
        {
            _generationUseCase.SetGenerationConfiguration(generationConfiguration);
            return _generationUseCase.WriteResolved(rules);
        }

        public GenerationWriteResult WritePreResolvedGeneration(IReadOnlyList<MeasurementRule> rules, GenerationConfiguration generationConfiguration)
        {
            _generationUseCase.SetGenerationConfiguration(generationConfiguration);
            return _generationUseCase.WritePreResolved(rules);
        }
    }

    public static class RecognitionProgress
    {
        [ThreadStatic]
        private static Action<int, string> _currentReporter;

        public static void SetReporter(Action<int, string> reporter)
        {
            _currentReporter = reporter;
        }

        public static void ClearReporter()
        {
            _currentReporter = null;
        }

        public static void Report(int percent, string message)
        {
            Trace.WriteLine($"[Host] Progress {percent}% - {message}");
            _currentReporter?.Invoke(percent, message);
        }
    }

    public sealed class RecognitionAndSyncResult
    {
        public RecognitionResult Recognition { get; set; }
        public SyncTemplateResult Remote { get; set; }
        public CachedTemplateRule Local { get; set; }
    }
}
