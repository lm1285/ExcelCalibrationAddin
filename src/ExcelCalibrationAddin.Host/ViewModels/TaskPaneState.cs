using System.Collections.Generic;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;

namespace ExcelCalibrationAddin.Host.ViewModels
{
    public sealed class TaskPaneState
    {
        public string WorkbookName { get; set; } = string.Empty;
        public string ExactFingerprint { get; set; } = string.Empty;
        public string RemoteTemplateId { get; set; } = string.Empty;
        public string RemoteTemplateName { get; set; } = string.Empty;
        public string RemoteStatus { get; set; } = "NotMatched";
        public string RemoteDetail { get; set; } = string.Empty;
        public string MatchStatus { get; set; } = "Pending";
        public string MatchStatusDetail { get; set; } = string.Empty;
        public string RecognitionStatusDetail { get; set; } = string.Empty;
        public string FingerprintStatusDetail { get; set; } = string.Empty;
        public string LocalMatchStatusDetail { get; set; } = string.Empty;
        public string RemoteMatchStatusDetail { get; set; } = string.Empty;
        public string DraftRuleStatusDetail { get; set; } = string.Empty;
        public TemplateLifecycleStatus? LocalTemplateStatus { get; set; }
        public string LocalTemplateStatusDetail { get; set; } = string.Empty;
        public double LocalMatchScore { get; set; }
        public bool IsCandidateMatch { get; set; }
        public GenerationConfiguration AppliedGenerationConfiguration { get; set; }
        public bool UsesTemplateGenerationConfiguration { get; set; }
        public TemplateFingerprint Fingerprint { get; set; }
        public bool IsFeatureBlocked { get; set; }
        public bool CanGenerate { get; set; }
        public IReadOnlyList<string> GenerationWarningMessages { get; set; } = new List<string>();
        public IReadOnlyList<RecognizedField> RecognizedFields { get; set; } = new List<RecognizedField>();
        public IReadOnlyList<TemplateRegionMapping> MappingItems { get; set; } = new List<TemplateRegionMapping>();
        public IReadOnlyList<MeasurementRule> DraftRules { get; set; } = new List<MeasurementRule>();
    }

    public sealed class SavedTemplateInfo
    {
        public string RemoteTemplateId { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public TemplateDirectoryMetadata DirectoryMetadata { get; set; } = new TemplateDirectoryMetadata();
        public string ExactFingerprint { get; set; } = string.Empty;
        public int RuleCount { get; set; }
        public System.DateTime UpdatedAt { get; set; }
        public TemplateLifecycleStatus Status { get; set; } = TemplateLifecycleStatus.Enabled;
        public TemplateSyncStatus LocalSyncStatus { get; set; } = TemplateSyncStatus.Synced;
        public string SyncError { get; set; } = string.Empty;
        public bool HasGenerationConfigurationOverride { get; set; }
        public bool HasRemoteConflict { get; set; }
    }

    public sealed class TemplateSaveResult
    {
        public bool SavedToRemote { get; set; }
        public bool SavedToLocal { get; set; }
        public TemplateSyncStatus LocalSyncStatus { get; set; } = TemplateSyncStatus.Synced;
        public string Message { get; set; } = string.Empty;
    }
}
