namespace ExcelCalibrationAddin.Core.Services
{
    public static class LocalRuleCacheSchema
    {
        public const string CreateRulesTable = @"
CREATE TABLE IF NOT EXISTS local_template_rules (
    id TEXT PRIMARY KEY,
    template_name TEXT NOT NULL,
    exact_fingerprint TEXT NOT NULL,
    fuzzy_fingerprint TEXT,
    fingerprint_json TEXT,
    directory_metadata_json TEXT,
    rule_json TEXT NOT NULL,
    generation_config_json TEXT,
    status INTEGER NOT NULL DEFAULT 0,
    local_sync_status INTEGER NOT NULL DEFAULT 0,
    sync_error TEXT,
    conflict_remote_json TEXT,
    remote_template_id TEXT,
    remote_version INTEGER,
    remote_updated_at TEXT,
    deleted_at TEXT,
    updated_at TEXT NOT NULL
);";

        public const string CreatePreferencesTable = @"
CREATE TABLE IF NOT EXISTS user_preferences (
    key TEXT PRIMARY KEY,
    value TEXT,
    updated_at TEXT NOT NULL
);";

        public const string CreateSampleDataTables = @"
CREATE TABLE IF NOT EXISTS SampleDataVersion (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    TemplateFingerprint TEXT NOT NULL,
    CreatedAt TEXT NOT NULL,
    Remark TEXT,
    SyncStatus INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS TemplateSampleData (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    VersionId INTEGER NOT NULL,
    CalibrationItemName TEXT NOT NULL,
    CalibrationItemKey TEXT,
    FOREIGN KEY (VersionId) REFERENCES SampleDataVersion(Id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS SampleDataPoint (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    SampleDataId INTEGER NOT NULL,
    PointIndex INTEGER NOT NULL,
    SourceRow INTEGER,
    SourceColumn INTEGER,
    StandardValue REAL NULL,
    MeasurementValues TEXT NOT NULL,
    DecimalPlaces INTEGER NOT NULL,
    FOREIGN KEY (SampleDataId) REFERENCES TemplateSampleData(Id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS IX_SampleDataVersion_Template ON SampleDataVersion(TemplateFingerprint, CreatedAt);
CREATE INDEX IF NOT EXISTS IX_TemplateSampleData_Version ON TemplateSampleData(VersionId);
CREATE INDEX IF NOT EXISTS IX_SampleDataPoint_Item ON SampleDataPoint(SampleDataId, PointIndex);";
    }
}
