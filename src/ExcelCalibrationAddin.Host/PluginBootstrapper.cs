using System;
using System.Net.Http;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Services;
using ExcelCalibrationAddin.Host.Templates;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host
{
    public sealed class PluginBootstrapper : IDisposable
    {
        public PluginBootstrapper(PluginConfiguration configuration)
        {
            Configuration = configuration;
            HttpClient = new HttpClient
            {
                BaseAddress = new Uri(configuration.Backend.BaseUrl),
                Timeout = TimeSpan.FromSeconds(10)
            };

            FingerprintBuilder = new TemplateFingerprintBuilder();
            FieldMatcher = new FieldMatcher();
            ParameterValueParser = new ParameterValueParser();
            NumberFormatInterpreter = new NumberFormatInterpreter();
            MeasurementRuleParameterResolver = new MeasurementRuleParameterResolver();
            TemplateSyncClient = new TemplateSyncClient(HttpClient, configuration.Backend.TemplateApiPrefix);
            LocalTemplateRuleCacheRepository = new LocalTemplateRuleCacheRepository(configuration.Cache.SqliteFile);
            LocalTemplateRuleCacheRepository.Initialize();
            MeasurementRuleDraftBuilder = new MeasurementRuleDraftBuilder(NumberFormatInterpreter);
            TemplateSaveService = new TemplateSaveService(TemplateSyncClient, LocalTemplateRuleCacheRepository);
            TemplatePendingUploadService = new TemplatePendingUploadService(TemplateSyncClient, LocalTemplateRuleCacheRepository);
        }

        public PluginConfiguration Configuration { get; }
        public HttpClient HttpClient { get; }
        public TemplateFingerprintBuilder FingerprintBuilder { get; }
        public FieldMatcher FieldMatcher { get; }
        public ParameterValueParser ParameterValueParser { get; }
        public NumberFormatInterpreter NumberFormatInterpreter { get; }
        public MeasurementRuleParameterResolver MeasurementRuleParameterResolver { get; }
        public TemplateSyncClient TemplateSyncClient { get; }
        public LocalTemplateRuleCacheRepository LocalTemplateRuleCacheRepository { get; }
        public MeasurementRuleDraftBuilder MeasurementRuleDraftBuilder { get; }
        public TemplateSaveService TemplateSaveService { get; }
        public TemplatePendingUploadService TemplatePendingUploadService { get; }

        public SampleDataUseCase CreateSampleDataUseCase(IWorkbookSnapshotProvider snapshotProvider)
        {
            return new SampleDataUseCase(snapshotProvider, LocalTemplateRuleCacheRepository);
        }

        public SyncTemplateUseCase CreateSyncTemplateUseCase()
        {
            return new SyncTemplateUseCase(TemplateSyncClient, LocalTemplateRuleCacheRepository);
        }

        public PluginWorkflowOrchestrator CreateWorkflowOrchestrator(IWorkbookSnapshotProvider snapshotProvider, IWorkbookWriter workbookWriter)
        {
            var recognitionUseCase = new TemplateRecognitionUseCase(snapshotProvider, FingerprintBuilder, FieldMatcher);
            var generationUseCase = new GenerateMeasurementUseCase(
                generationConfiguration => new MeasurementValueGenerator(generationConfiguration),
                Configuration.Generation,
                workbookWriter,
                snapshotProvider,
                MeasurementRuleParameterResolver);
            var syncTemplateUseCase = CreateSyncTemplateUseCase();
            return new PluginWorkflowOrchestrator(recognitionUseCase, generationUseCase, syncTemplateUseCase, LocalTemplateRuleCacheRepository);
        }

        public void Dispose()
        {
            HttpClient.Dispose();
        }
    }
}
