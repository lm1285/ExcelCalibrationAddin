using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.ViewModels;

namespace ExcelCalibrationAddin.Host.Templates
{
    public sealed class TemplateSaveService
    {
        private readonly TemplateSyncClient _syncClient;
        private readonly LocalTemplateRuleCacheRepository _cacheRepository;

        public TemplateSaveService(
            TemplateSyncClient syncClient,
            LocalTemplateRuleCacheRepository cacheRepository)
        {
            _syncClient = syncClient ?? throw new ArgumentNullException(nameof(syncClient));
            _cacheRepository = cacheRepository ?? throw new ArgumentNullException(nameof(cacheRepository));
        }

        public TemplateSaveResult Save(
            string templateName,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration,
            bool createNew,
            TemplateDirectoryMetadata directoryMetadata = null,
            string targetRemoteTemplateId = null)
        {
            return SaveAsync(
                    templateName,
                    fingerprint,
                    rules,
                    generationConfiguration,
                    createNew,
                    directoryMetadata,
                    targetRemoteTemplateId)
                .GetAwaiter()
                .GetResult();
        }

        public async System.Threading.Tasks.Task<TemplateSaveResult> SaveAsync(
            string templateName,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration,
            bool createNew,
            TemplateDirectoryMetadata directoryMetadata = null,
            string targetRemoteTemplateId = null)
        {
            Trace.WriteLine($"[Host] TemplateSave enter. Name={templateName}, Rules={rules?.Count ?? 0}");
            _cacheRepository.ValidateTemplateForSave(rules);
            var persistentRules = TemplateRulePersistencePreparer.Prepare(rules);
            try
            {
                var existing = string.IsNullOrWhiteSpace(targetRemoteTemplateId)
                    ? _cacheRepository.FindByExactFingerprint(fingerprint?.ExactFingerprint ?? string.Empty)
                    : _cacheRepository.FindByRemoteTemplateId(targetRemoteTemplateId);
                var remoteResponse = await _syncClient.SaveTemplateAsync(
                        templateName,
                        fingerprint,
                        persistentRules,
                        generationConfiguration,
                        createNew ? null : existing?.RemoteTemplateId,
                        createNew,
                        directoryMetadata)
                    .ConfigureAwait(false);

                var acceptedTemplate = RemoteTemplateSaveResponseParser.Parse(
                    remoteResponse,
                    templateName,
                    fingerprint,
                    persistentRules,
                    generationConfiguration);

                // The submitted fingerprint identifies the workbook that this save updates.
                // A remote service may normalize or omit its echoed fingerprint, but that must
                // not change the local key used by subsequent template-library matching.
                acceptedTemplate.Fingerprint = fingerprint;
                acceptedTemplate.Rules = TemplateRulePersistencePreparer.MergeAcceptedRules(
                    acceptedTemplate.Rules,
                    persistentRules);
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
                    directoryMetadata ?? acceptedTemplate.DirectoryMetadata,
                    createNew ? null : existing?.RemoteTemplateId);

                Trace.WriteLine($"[Host] TemplateSave remote success. ResponseLength={remoteResponse?.Length ?? 0}");
                return new TemplateSaveResult
                {
                    SavedToRemote = true,
                    SavedToLocal = true,
                    LocalSyncStatus = TemplateSyncStatus.Synced,
                    Message = "模板已保存并同步到模板库。"
                };
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Host] TemplateSave remote failed, saving pending upload. Error={ex.Message}");
                _cacheRepository.SaveTemplate(
                    templateName,
                    fingerprint,
                    persistentRules,
                    generationConfiguration,
                    createNew,
                    TemplateSyncStatus.PendingUpload,
                    ex.Message,
                    directoryMetadata,
                    targetRemoteTemplateId);

                return new TemplateSaveResult
                {
                    SavedToRemote = false,
                    SavedToLocal = true,
                    LocalSyncStatus = TemplateSyncStatus.PendingUpload,
                    Message = "远端模板库不可用，模板已本地暂存并标记为待上传。"
                };
            }
        }
    }
}
