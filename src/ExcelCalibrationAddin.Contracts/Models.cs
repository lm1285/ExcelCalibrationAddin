using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelCalibrationAddin.Contracts
{
    public enum ErrorType
    {
        Absolute,
        Relative,
        Referenced
    }

    public enum TechnicalRequirementOperator
    {
        None,
        PlusMinus,
        LessThan,
        LessThanOrEqual,
        GreaterThan,
        GreaterThanOrEqual
    }

    public enum FillMode
    {
        Row,
        Column,
        Block
    }

    public enum DistributionMode
    {
        Normal,
        TruncatedNormal = Normal,
        Triangular,
        Uniform
    }

    public enum TemplateLifecycleStatus
    {
        Enabled = 0,
        Obsolete = 1,
        Disabled = 2
    }

    public enum TemplateSyncStatus
    {
        Synced = 0,
        PendingUpload = 1,
        Conflict = 2,
        SyncFailed = 3
    }

    public sealed class CellRange
    {
        public string SheetName { get; set; } = string.Empty;
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public int EndRow { get; set; }
        public int EndColumn { get; set; }

        public override string ToString()
        {
            return $"{SheetName}:{StartRow},{StartColumn}-{EndRow},{EndColumn}";
        }
    }

    public sealed class CellAddress
    {
        public int Row { get; set; }
        public int Column { get; set; }
    }

    public sealed class CellMeta
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Text { get; set; } = string.Empty;
        public string RawValueText { get; set; } = string.Empty;
        public string DisplayText { get; set; } = string.Empty;
        public string NumberFormat { get; set; } = string.Empty;
        public bool IsMerged { get; set; }
        public CellRange MergeRange { get; set; }
        public bool HasBorder { get; set; }
        public string BackgroundColor { get; set; } = string.Empty;
        public string Formula { get; set; } = string.Empty;
        public string FormulaR1C1 { get; set; } = string.Empty;
        public double RowHeight { get; set; }
        public double ColumnWidth { get; set; }
    }

    public sealed class HeaderPath
    {
        public int Column { get; set; }
        public List<string> Levels { get; set; } = new List<string>();
        public string FullText => string.Join("/", Levels);
    }

    public sealed class SheetSnapshot
    {
        public string Name { get; set; } = string.Empty;
        public string UsedRangeShape { get; set; } = string.Empty;
        public List<CellMeta> Cells { get; set; } = new List<CellMeta>();
        public List<HeaderPath> Headers { get; set; } = new List<HeaderPath>();
    }

    public sealed class WorkbookSnapshot
    {
        public string WorkbookName { get; set; } = string.Empty;
        public List<SheetSnapshot> Sheets { get; set; } = new List<SheetSnapshot>();
    }

    public sealed class FormatRule
    {
        public int? DecimalPlaces { get; set; }
        public bool IsScientificNotation { get; set; }
        public bool IsPercentage { get; set; }
        public string UnitSuffix { get; set; } = string.Empty;
        public string RawNumberFormat { get; set; } = string.Empty;
    }

    public sealed class ParameterSource
    {
        public string Name { get; set; } = string.Empty;
        public CellRange Range { get; set; }
        public string ValuePattern { get; set; } = string.Empty;
    }

    public sealed class MeasurementRule
    {
        public string FieldName { get; set; } = string.Empty;
        public string FieldAlias { get; set; } = string.Empty;
        public CellRange TargetRange { get; set; }
        public ErrorType ErrorType { get; set; }
        public FillMode FillMode { get; set; }
        public DistributionMode DistributionMode { get; set; } = DistributionMode.TruncatedNormal;
        public FormatRule FormatRule { get; set; } = new FormatRule();
        public ParameterSource SetpointSource { get; set; }
        public ParameterSource StandardValueSource { get; set; }
        public ParameterSource AverageSource { get; set; }
        public ParameterSource ErrorSource { get; set; }
        public ParameterSource MpeSource { get; set; }
        public ParameterSource RangeSource { get; set; }
        public ParameterSource UncertaintySource { get; set; }
        public ParameterSource ResultSource { get; set; }
        public double? FixedStandardValue { get; set; }
        public List<ManualStandardValue> ManualStandardValues { get; set; } = new List<ManualStandardValue>();
        public double? FixedMpe { get; set; }
        public double? FixedNegativeTolerance { get; set; }
        public double? FixedPositiveTolerance { get; set; }
        public TechnicalRequirementOperator RequirementOperator { get; set; }
        public double? FixedReferenceRange { get; set; }
        public double? MeasurementLowerBound { get; set; }
        public double? MeasurementUpperBound { get; set; }
        public List<CellAddress> WritableCells { get; set; } = new List<CellAddress>();
        public List<MeasurementRowMapping> RowMappings { get; set; } = new List<MeasurementRowMapping>();
        public int GroupSize { get; set; } = 1;
        public bool IsEnabled { get; set; } = true;
        public bool PositiveDirectionOnly { get; set; }
        public bool NegativeDirectionOnly { get; set; }
        public MeasurementGenerationCoefficientOverride GenerationCoefficientOverride { get; set; }
        public ErrorFormulaInfo ErrorFormula { get; set; }
        public TemplateFieldDefinition TemplateDefinition { get; set; }
    }

    public sealed class ManualStandardValue
    {
        // One-based ordinal of the measurement point within the rule's writable rows.
        public int PointIndex { get; set; }
        public double? Value { get; set; }
    }

    public sealed class MeasurementRowMapping
    {
        public int Row { get; set; }
        public CellRange SetpointValueRange { get; set; }
        public CellRange StandardValueRange { get; set; }
        public List<CellAddress> MeasurementCells { get; set; } = new List<CellAddress>();
        public CellRange AverageRange { get; set; }
        public CellRange ErrorRange { get; set; }
        public CellRange TechnicalRequirementRange { get; set; }
        public CellRange RangeValueRange { get; set; }
        public CellRange UncertaintyRange { get; set; }
        public CellRange ResultRange { get; set; }
        public bool IsComplete { get; set; }
        public string StatusMessage { get; set; } = string.Empty;
    }

    public sealed class MeasurementGenerationOverride
    {
        public string FieldName { get; set; } = string.Empty;
        public MeasurementGenerationCoefficientOverride CoefficientOverride { get; set; }
        public int? DecimalPlaces { get; set; }
        public double? AlarmValue { get; set; }
    }

    public sealed class MeasurementGenerationCoefficientOverride
    {
        public double? NegativeMinimumCoefficient { get; set; }
        public double? NegativeMaximumCoefficient { get; set; }
        public double? PositiveMinimumCoefficient { get; set; }
        public double? PositiveMaximumCoefficient { get; set; }
        public double? AbsoluteMinimumCoefficient { get; set; }
        public double? AbsoluteMaximumCoefficient { get; set; }
    }

    public sealed class ErrorFormulaInfo
    {
        public bool HasFormula { get; set; }
        public string Formula { get; set; } = string.Empty;
        public bool ReferencesMeasurement { get; set; }
        public bool ReferencesStandardValue { get; set; }
        public bool ReferencesAverage { get; set; }
        public bool AverageFormulaResolved { get; set; }
        public string AverageFormula { get; set; } = string.Empty;
        public bool TechnicalRequirementFormulaResolved { get; set; }
        public string TechnicalRequirementFormula { get; set; } = string.Empty;
        public bool UncertaintyFormulaResolved { get; set; }
        public string UncertaintyFormula { get; set; } = string.Empty;
        public bool ResultFormulaResolved { get; set; }
        public string ResultFormula { get; set; } = string.Empty;
        public ErrorFormulaScale Scale { get; set; } = ErrorFormulaScale.Absolute;
        public bool FormulaMultipliesBy100 { get; set; }
        public bool FormulaDividesByReferenceRange { get; set; }
    }

    public static class ErrorFormulaClassifier
    {
        public static bool IsMaximumError(MeasurementRule rule)
        {
            if (rule == null)
            {
                return false;
            }

            var name = string.IsNullOrWhiteSpace(rule.FieldAlias) ? rule.FieldName : rule.FieldAlias;
            if ((name ?? string.Empty).IndexOf("稳定性", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            return IsMaximumError(rule.ErrorFormula);
        }

        public static bool IsMaximumError(ErrorFormulaInfo info)
        {
            if (info == null)
            {
                return false;
            }

            var normalized = NormalizeFormula(info.Formula);
            if (normalized.Contains("MAX(") &&
                normalized.Contains("MIN(") &&
                info.ReferencesMeasurement &&
                !info.ReferencesStandardValue &&
                !info.ReferencesAverage)
            {
                return false;
            }

            return ClassifyMaximumError(normalized);
        }

        public static bool IsMaximumError(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return false;
            }

            return ClassifyMaximumError(NormalizeFormula(formula));
        }

        private static bool ClassifyMaximumError(string normalized)
        {
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            if (normalized.Contains("MAX(") && normalized.Contains("MIN(") && !normalized.Contains("MAXIFS("))
            {
                return true;
            }

            return normalized.Contains("MAX(") || normalized.Contains("MAXIFS(");
        }

        private static string NormalizeFormula(string formula)
        {
            return new string((formula ?? string.Empty)
                .Where(character => !char.IsWhiteSpace(character))
                .ToArray())
                .ToUpperInvariant();
        }
    }

    public enum ErrorFormulaScale
    {
        Absolute,
        RelativeToStandardValue,
        RelativeToReferenceRange
    }

    public sealed class TemplateFingerprint
    {
        public string ExactFingerprint { get; set; } = string.Empty;
        public string FuzzyFingerprint { get; set; } = string.Empty;
        public string StructureSignature { get; set; } = string.Empty;
        public string Summary { get; set; } = string.Empty;
        public List<string> SheetNames { get; set; } = new List<string>();
        public string Title { get; set; } = string.Empty;
        public List<string> HeaderTexts { get; set; } = new List<string>();
    }

    public sealed class TemplateDirectoryMetadata
    {
        public string MeasurementDomain { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public string VariantName { get; set; } = string.Empty;
        public string TemplateCode { get; set; } = string.Empty;
    }

    public sealed class RecognizedField
    {
        public string Alias { get; set; } = string.Empty;
        public CellRange Range { get; set; }
        public double Score { get; set; }
        public string Reason { get; set; } = string.Empty;
    }

    public sealed class TemplateRegionMapping
    {
        public string ProjectName { get; set; } = string.Empty;
        public CellRange SectionRange { get; set; }
        public CellRange SetpointValueRange { get; set; }
        public CellRange StandardValueRange { get; set; }
        public CellRange MeasurementValueRange { get; set; }
        public CellRange AverageValueRange { get; set; }
        public CellRange ErrorValueRange { get; set; }
        public CellRange TechnicalRequirementRange { get; set; }
        public CellRange UncertaintyRange { get; set; }
        public CellRange RangeValueRange { get; set; }
        public CellRange ResultRange { get; set; }
        public string Notes { get; set; } = string.Empty;
    }

    public sealed class MeasurementGenerationInput
    {
        public double StandardValue { get; set; }
        public double Mpe { get; set; }
        public double? NegativeTolerance { get; set; }
        public double? PositiveTolerance { get; set; }
        public TechnicalRequirementOperator RequirementOperator { get; set; }
        public double? ReferenceRange { get; set; }
        public ErrorType ErrorType { get; set; }
        public DistributionMode DistributionMode { get; set; } = DistributionMode.TruncatedNormal;
        public int ValueCount { get; set; } = 1;
        public int DecimalPlaces { get; set; } = 2;
        public List<int> DecimalPlacesByValue { get; set; } = new List<int>();
        public bool ForcePositiveDirection { get; set; }
        public bool ForceNegativeDirection { get; set; }
        public bool UseSameDeviationDirection { get; set; } = true;
        public bool UseIndependentDeviationControl { get; set; } = true;
        public double? AnchorError { get; set; }
        public MeasurementGenerationCoefficientOverride CoefficientOverride { get; set; }
        public double? MeasurementLowerBound { get; set; }
        public double? MeasurementUpperBound { get; set; }
    }

    public sealed class MeasurementGenerationResult
    {
        public List<double> RawValues { get; set; } = new List<double>();
        public List<string> DisplayValues { get; set; } = new List<string>();
        public double LowerBound { get; set; }
        public double UpperBound { get; set; }
        public int Direction { get; set; }
    }

}
