using System;
using System.Diagnostics;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;

namespace ExcelCalibrationAddin.Host.Templates
{
    public sealed class TemplatePendingUploadService
    {
        private readonly TemplateSyncClient _syncClient;
        private readonly LocalTemplateRuleCacheRepository _cacheRepository;

        public TemplatePendingUploadService(
            TemplateSyncClient syncClient,
            LocalTemplateRuleCacheRepository cacheRepository)
        {
            _syncClient = syncClient ?? throw new ArgumentNullException(nameof(syncClient));
            _cacheRepository = cacheRepository ?? throw new ArgumentNullException(nameof(cacheRepository));
        }

        public int UploadPendingTemplates()
        {
            var pendingTemplates = _cacheRepository.ListTemplatesBySyncStatus(
                TemplateSyncStatus.PendingUpload,
                TemplateSyncStatus.SyncFailed);
            var uploadedCount = 0;

            foreach (var template in pendingTemplates)
            {
                try
                {
                    var remoteResponse = _syncClient
                        .SaveTemplateAsync(
                             template.TemplateName,
                             template.Fingerprint,
                             template.Rules,
                             template.GenerationConfiguration,
                             template.RemoteTemplateId,
                             IsLocallyCreatedTemplateId(template.RemoteTemplateId),
                             template.DirectoryMetadata)
                        .GetAwaiter()
                        .GetResult();
                    var acceptedTemplate = RemoteTemplateSaveResponseParser.Parse(
                        remoteResponse,
                        template.TemplateName,
                        template.Fingerprint,
                        template.Rules,
                        template.GenerationConfiguration);

                    _cacheRepository.SaveAcceptedRemoteTemplate(
                        acceptedTemplate.TemplateId,
                        acceptedTemplate.TemplateName,
                        acceptedTemplate.Version,
                        acceptedTemplate.RemoteUpdatedAt,
                        acceptedTemplate.DeletedAt,
                        acceptedTemplate.Status,
                        acceptedTemplate.Fingerprint,
                        acceptedTemplate.Rules,
                        acceptedTemplate.GenerationConfiguration,
                        template.DirectoryMetadata,
                        template.RemoteTemplateId);
                    uploadedCount++;
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[Host] Upload pending template failed. Name={template.TemplateName}, Error={ex.Message}");
                    _cacheRepository.UpdateTemplateSyncStatus(
                        template.ExactFingerprint,
                        TemplateSyncStatus.SyncFailed,
                        ex.Message);
                }
            }

            return uploadedCount;
        }

        private static bool IsLocallyCreatedTemplateId(string templateId)
        {
            return string.IsNullOrWhiteSpace(templateId) ||
                templateId.StartsWith("template:", StringComparison.OrdinalIgnoreCase);
        }
    }
}
