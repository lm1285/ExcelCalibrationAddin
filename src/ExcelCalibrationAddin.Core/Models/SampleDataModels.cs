using System;
using System.Collections.Generic;

namespace ExcelCalibrationAddin.Core.Models
{
    public sealed class SampleDataVersion
    {
        public long Id { get; set; }
        public string TemplateFingerprint { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; }
        public string Remark { get; set; } = string.Empty;
        public int SyncStatus { get; set; }
        public List<TemplateSampleData> Items { get; set; } = new List<TemplateSampleData>();
        public int ItemCount => Items?.Count ?? 0;
    }

    public sealed class TemplateSampleData
    {
        public long Id { get; set; }
        public long VersionId { get; set; }
        public string CalibrationItemName { get; set; } = string.Empty;
        public string CalibrationItemKey { get; set; } = string.Empty;
        public List<SampleDataPoint> Points { get; set; } = new List<SampleDataPoint>();
    }

    public sealed class SampleDataPoint
    {
        public string CalibrationItemName { get; set; } = string.Empty;
        public long Id { get; set; }
        public long SampleDataId { get; set; }
        public int PointIndex { get; set; }
        public int SourceRow { get; set; }
        public int SourceColumn { get; set; }
        public double? StandardValue { get; set; }
        public List<double> MeasurementValues { get; set; } = new List<double>();
        public int DecimalPlaces { get; set; }
    }
}
