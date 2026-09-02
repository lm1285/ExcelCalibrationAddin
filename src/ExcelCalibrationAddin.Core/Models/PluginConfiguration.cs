namespace ExcelCalibrationAddin.Core.Models
{
    public sealed class PluginConfiguration
    {
        public BackendConfiguration Backend { get; set; } = new BackendConfiguration();
        public CacheConfiguration Cache { get; set; } = new CacheConfiguration();
        public GenerationConfiguration Generation { get; set; } = new GenerationConfiguration();
        public AutomationConfiguration Automation { get; set; } = new AutomationConfiguration();
    }

    public sealed class BackendConfiguration
    {
        public string BaseUrl { get; set; } = "https://www.wzglpt.top";
        public string TemplateApiPrefix { get; set; } = "/api/excel-templates";
        public string AuthorizationToken { get; set; } = string.Empty;
    }

    public sealed class AutomationConfiguration
    {
        public bool Enabled { get; set; } = true;
        public int Port { get; set; } = 30771;
        public int InternalPort { get; set; } = 30772;
        public string Token { get; set; } = string.Empty;
    }

    public sealed class CacheConfiguration
    {
        public string SqliteFile { get; set; } = "%LOCALAPPDATA%\\ExcelCalibrationAddin\\cache.db";
    }

    public sealed class GenerationConfiguration
    {
        public string GenerateShortcutKey { get; set; } = "F6";
        public string DefaultDistribution { get; set; } = "Normal";
        public string ResultCalculationMethod { get; set; } = "FormulaBackCalculation";
        public string StandardValueReference { get; set; } = "RecognizedStandardValueRange";
        public bool UseSameDeviationDirection { get; set; } = true;
        public bool UseIndependentDeviationControl { get; set; } = true;
        public double UnifiedErrorMinimumCoefficient { get; set; } = 0.2;
        public double UnifiedErrorMaximumCoefficient { get; set; } = 0.8;
        public double PositiveErrorMinimumCoefficient { get; set; } = 0.2;
        public double PositiveErrorMaximumCoefficient { get; set; } = 0.8;
        public double NegativeErrorMinimumCoefficient { get; set; } = 0.2;
        public double NegativeErrorMaximumCoefficient { get; set; } = 0.8;
        public double AbsoluteErrorMinimumCoefficient { get; set; } = 0.2;
        public double AbsoluteErrorMaximumCoefficient { get; set; } = 0.8;
        public double MinimumRequirementMinimumCoefficient { get; set; } = 1.05;
        public double MinimumRequirementMaximumCoefficient { get; set; } = 1.30;
        // Kept for reading legacy configuration files. Resolution is always derived from the template.
        public bool UseDecimalPlacesForResolution { get; set; } = true;
        public bool ShouldSerializeUseDecimalPlacesForResolution() => false;
        public double MeasurementGroupMinimumFluctuationCoefficient { get; set; } = 0.01;
        public double MeasurementGroupMaximumFluctuationCoefficient { get; set; } = 0.06;
        public double ResultGroupMinimumFluctuationCoefficient { get; set; }
        public double ResultGroupMaximumFluctuationCoefficient { get; set; } = 0.20;
        public double ResponseTimeThresholdSeconds { get; set; } = 180;
        public double ResponseTimeBelowThresholdMaximumDifferenceSeconds { get; set; } = 5;
        public double ResponseTimeAboveThresholdMaximumDifferenceSeconds { get; set; } = 7;

        // Legacy JSON alias for the former single upper bound.
        public double ResultGroupFluctuationCoefficient
        {
            get => ResultGroupMaximumFluctuationCoefficient;
            set => ResultGroupMaximumFluctuationCoefficient = value;
        }

        public bool ShouldSerializeResultGroupFluctuationCoefficient() => false;

        public double MeasurementGroupFluctuationCoefficient
        {
            get => MeasurementGroupMaximumFluctuationCoefficient;
            set => MeasurementGroupMaximumFluctuationCoefficient = value;
        }

        public bool ShouldSerializeMeasurementGroupFluctuationCoefficient() => false;

        public double SameStandardValueFluctuationCoefficient
        {
            get => MeasurementGroupMaximumFluctuationCoefficient;
            set => MeasurementGroupMaximumFluctuationCoefficient = value;
        }

        public bool ShouldSerializeSameStandardValueFluctuationCoefficient() => false;

        public double CrossStandardValueFluctuationCoefficient
        {
            get => ResultGroupMaximumFluctuationCoefficient;
            set => ResultGroupMaximumFluctuationCoefficient = value;
        }

        public bool ShouldSerializeCrossStandardValueFluctuationCoefficient() => false;
    }
}
