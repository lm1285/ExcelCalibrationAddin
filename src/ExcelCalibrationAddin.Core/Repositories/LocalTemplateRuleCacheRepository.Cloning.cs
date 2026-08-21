using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Repositories
{
    public sealed partial class LocalTemplateRuleCacheRepository
    {
        private static MeasurementRule CloneTemplateRule(MeasurementRule rule)
        {
            if (rule == null)
            {
                return null;
            }

            return new MeasurementRule
            {
                FieldName = rule.FieldName,
                FieldAlias = rule.FieldAlias,
                TargetRange = CloneRange(rule.TargetRange),
                ErrorType = rule.ErrorType,
                FillMode = rule.FillMode,
                DistributionMode = rule.DistributionMode,
                FormatRule = CloneFormatRule(rule.FormatRule),
                SetpointSource = CloneParameterSource(rule.SetpointSource),
                StandardValueSource = CloneParameterSource(rule.StandardValueSource),
                AverageSource = CloneParameterSource(rule.AverageSource),
                ErrorSource = CloneParameterSource(rule.ErrorSource),
                MpeSource = CloneParameterSource(rule.MpeSource),
                RangeSource = CloneParameterSource(rule.RangeSource),
                UncertaintySource = CloneParameterSource(rule.UncertaintySource),
                ResultSource = CloneParameterSource(rule.ResultSource),
                FixedStandardValue = rule.FixedStandardValue,
                ManualStandardValues = CloneManualStandardValues(rule.ManualStandardValues),
                FixedMpe = rule.FixedMpe,
                FixedNegativeTolerance = rule.FixedNegativeTolerance,
                FixedPositiveTolerance = rule.FixedPositiveTolerance,
                RequirementOperator = rule.RequirementOperator,
                FixedReferenceRange = rule.FixedReferenceRange,
                MeasurementLowerBound = rule.MeasurementLowerBound,
                MeasurementUpperBound = rule.MeasurementUpperBound,
                WritableCells = CloneCellAddresses(rule.WritableCells),
                RowMappings = CloneRowMappings(rule.RowMappings),
                GroupSize = rule.GroupSize,
                IsEnabled = rule.IsEnabled,
                PositiveDirectionOnly = rule.PositiveDirectionOnly,
                NegativeDirectionOnly = rule.NegativeDirectionOnly,
                GenerationCoefficientOverride = CloneCoefficientOverride(rule.GenerationCoefficientOverride),
                ErrorFormula = CloneErrorFormula(rule.ErrorFormula),
                TemplateDefinition = TemplateDefinitionCloner.Clone(rule.TemplateDefinition)
            };
        }

        private static FormatRule CloneFormatRule(FormatRule rule)
        {
            if (rule == null)
            {
                return null;
            }

            return new FormatRule
            {
                DecimalPlaces = rule.DecimalPlaces,
                IsScientificNotation = rule.IsScientificNotation,
                IsPercentage = rule.IsPercentage,
                UnitSuffix = rule.UnitSuffix,
                RawNumberFormat = rule.RawNumberFormat
            };
        }

        private static ParameterSource CloneParameterSource(ParameterSource source)
        {
            if (source == null)
            {
                return null;
            }

            return new ParameterSource
            {
                Name = source.Name,
                Range = CloneRange(source.Range),
                ValuePattern = source.ValuePattern
            };
        }

        private static CellRange CloneRange(CellRange range)
        {
            if (range == null)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                StartColumn = range.StartColumn,
                EndRow = range.EndRow,
                EndColumn = range.EndColumn
            };
        }

        private static List<CellAddress> CloneCellAddresses(IEnumerable<CellAddress> cells)
        {
            return (cells ?? Enumerable.Empty<CellAddress>())
                .Where(cell => cell != null && cell.Row > 0 && cell.Column > 0)
                .Select(cell => new CellAddress { Row = cell.Row, Column = cell.Column })
                .ToList();
        }

        private static List<ManualStandardValue> CloneManualStandardValues(IEnumerable<ManualStandardValue> values)
        {
            return (values ?? Enumerable.Empty<ManualStandardValue>())
                .Where(item => item != null)
                .Select(item => new ManualStandardValue { PointIndex = item.PointIndex, Value = item.Value })
                .ToList();
        }

        private static List<MeasurementRowMapping> CloneRowMappings(IEnumerable<MeasurementRowMapping> mappings)
        {
            return (mappings ?? Enumerable.Empty<MeasurementRowMapping>())
                .Where(item => item != null)
                .Select(item => new MeasurementRowMapping
                {
                    Row = item.Row,
                    SetpointValueRange = CloneRange(item.SetpointValueRange),
                    StandardValueRange = CloneRange(item.StandardValueRange),
                    MeasurementCells = CloneCellAddresses(item.MeasurementCells),
                    AverageRange = CloneRange(item.AverageRange),
                    ErrorRange = CloneRange(item.ErrorRange),
                    TechnicalRequirementRange = CloneRange(item.TechnicalRequirementRange),
                    RangeValueRange = CloneRange(item.RangeValueRange),
                    UncertaintyRange = CloneRange(item.UncertaintyRange),
                    ResultRange = CloneRange(item.ResultRange),
                    IsComplete = item.IsComplete,
                    StatusMessage = item.StatusMessage
                })
                .ToList();
        }

        private static MeasurementGenerationCoefficientOverride CloneCoefficientOverride(MeasurementGenerationCoefficientOverride value)
        {
            if (value == null)
            {
                return null;
            }

            return new MeasurementGenerationCoefficientOverride
            {
                NegativeMinimumCoefficient = value.NegativeMinimumCoefficient,
                NegativeMaximumCoefficient = value.NegativeMaximumCoefficient,
                PositiveMinimumCoefficient = value.PositiveMinimumCoefficient,
                PositiveMaximumCoefficient = value.PositiveMaximumCoefficient,
                AbsoluteMinimumCoefficient = value.AbsoluteMinimumCoefficient,
                AbsoluteMaximumCoefficient = value.AbsoluteMaximumCoefficient
            };
        }

        private static ErrorFormulaInfo CloneErrorFormula(ErrorFormulaInfo info)
        {
            if (info == null)
            {
                return null;
            }

            return new ErrorFormulaInfo
            {
                HasFormula = info.HasFormula,
                Formula = info.Formula,
                ReferencesMeasurement = info.ReferencesMeasurement,
                ReferencesStandardValue = info.ReferencesStandardValue,
                ReferencesAverage = info.ReferencesAverage,
                AverageFormulaResolved = info.AverageFormulaResolved,
                AverageFormula = info.AverageFormula,
                TechnicalRequirementFormulaResolved = info.TechnicalRequirementFormulaResolved,
                TechnicalRequirementFormula = info.TechnicalRequirementFormula,
                UncertaintyFormulaResolved = info.UncertaintyFormulaResolved,
                UncertaintyFormula = info.UncertaintyFormula,
                ResultFormulaResolved = info.ResultFormulaResolved,
                ResultFormula = info.ResultFormula,
                Scale = info.Scale,
                FormulaMultipliesBy100 = info.FormulaMultipliesBy100,
                FormulaDividesByReferenceRange = info.FormulaDividesByReferenceRange
            };
        }
    }
}
