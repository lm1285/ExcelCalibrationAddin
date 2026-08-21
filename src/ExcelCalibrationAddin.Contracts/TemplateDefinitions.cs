using System.Collections.Generic;
using System.Linq;

namespace ExcelCalibrationAddin.Contracts
{
    public enum TemplateRegionRole
    {
        Unknown,
        StandardValue,
        MeasurementValue,
        AverageValue,
        ErrorValue,
        TechnicalRequirement,
        RangeValue,
        Uncertainty,
        Result,
        SetpointValue
    }

    public enum TemplateUnitSource
    {
        None,
        CellDisplay,
        CellFormat,
        RegionHeader,
        FieldHeader,
        FormulaBranch
    }

    public sealed class TemplateFieldDefinition
    {
        public int SchemaVersion { get; set; } = 1;
        public string ProjectName { get; set; } = string.Empty;
        public CellRange SectionRange { get; set; }
        public List<TemplateHeaderDefinition> Headers { get; set; } = new List<TemplateHeaderDefinition>();
        public List<TemplateRegionDefinition> Regions { get; set; } = new List<TemplateRegionDefinition>();
    }

    public sealed class TemplateHeaderDefinition
    {
        public int Level { get; set; }
        public CellRange Range { get; set; }
        public string Text { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public string NumberFormat { get; set; } = string.Empty;
    }

    public sealed class TemplateRegionDefinition
    {
        public TemplateRegionRole Role { get; set; }
        public CellRange Range { get; set; }
        public CellRange HeaderRange { get; set; }
        public List<string> HeaderPath { get; set; } = new List<string>();
        public string Unit { get; set; } = string.Empty;
        public List<string> Units { get; set; } = new List<string>();
        public TemplateUnitSource UnitSource { get; set; }
        public string NumberFormat { get; set; } = string.Empty;
        public CellRange OperatorRange { get; set; }
        public CellRange ValueRange { get; set; }
        public List<TemplateRequirementValue> RequirementValues { get; set; } = new List<TemplateRequirementValue>();
        public TemplateFormulaDefinition Formula { get; set; }
        public List<TemplateFormulaDefinition> FormulaVariants { get; set; } = new List<TemplateFormulaDefinition>();
    }

    public sealed class TemplateRequirementValue
    {
        public CellRange Range { get; set; }
        public TechnicalRequirementOperator Operator { get; set; }
        public double? Value { get; set; }
        public double? NegativeValue { get; set; }
        public double? PositiveValue { get; set; }
        public string Unit { get; set; } = string.Empty;
        public string DisplayText { get; set; } = string.Empty;
        public string NumberFormat { get; set; } = string.Empty;
    }

    public sealed class TemplateFormulaDefinition
    {
        public string Formula { get; set; } = string.Empty;
        public string FormulaR1C1 { get; set; } = string.Empty;
        public bool IsConditional { get; set; }
        public bool IsFullyParsed { get; set; }
        public List<TemplateFormulaReference> References { get; set; } = new List<TemplateFormulaReference>();
        public List<TemplateFormulaBranch> Branches { get; set; } = new List<TemplateFormulaBranch>();
    }

    public sealed class TemplateFormulaReference
    {
        public string Token { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public CellRange Range { get; set; }
        public TemplateRegionRole Role { get; set; }
    }

    public sealed class TemplateFormulaBranch
    {
        public string Condition { get; set; } = string.Empty;
        public string ValueExpression { get; set; } = string.Empty;
        public string DisplayText { get; set; } = string.Empty;
        public double? NumericValue { get; set; }
        public TechnicalRequirementOperator Operator { get; set; }
        public string Unit { get; set; } = string.Empty;
    }

    public static class TemplateDefinitionCloner
    {
        public static TemplateFieldDefinition Clone(TemplateFieldDefinition source)
        {
            if (source == null)
            {
                return null;
            }

            return new TemplateFieldDefinition
            {
                SchemaVersion = source.SchemaVersion,
                ProjectName = source.ProjectName,
                SectionRange = CloneRange(source.SectionRange),
                Headers = (source.Headers ?? new List<TemplateHeaderDefinition>())
                    .Where(item => item != null)
                    .Select(CloneHeader)
                    .ToList(),
                Regions = (source.Regions ?? new List<TemplateRegionDefinition>())
                    .Where(item => item != null)
                    .Select(CloneRegion)
                    .ToList()
            };
        }

        private static TemplateHeaderDefinition CloneHeader(TemplateHeaderDefinition source)
        {
            return new TemplateHeaderDefinition
            {
                Level = source.Level,
                Range = CloneRange(source.Range),
                Text = source.Text,
                Unit = source.Unit,
                NumberFormat = source.NumberFormat
            };
        }

        private static TemplateRegionDefinition CloneRegion(TemplateRegionDefinition source)
        {
            return new TemplateRegionDefinition
            {
                Role = source.Role,
                Range = CloneRange(source.Range),
                HeaderRange = CloneRange(source.HeaderRange),
                HeaderPath = (source.HeaderPath ?? new List<string>()).ToList(),
                Unit = source.Unit,
                Units = (source.Units ?? new List<string>()).ToList(),
                UnitSource = source.UnitSource,
                NumberFormat = source.NumberFormat,
                OperatorRange = CloneRange(source.OperatorRange),
                ValueRange = CloneRange(source.ValueRange),
                RequirementValues = (source.RequirementValues ?? new List<TemplateRequirementValue>())
                    .Where(item => item != null)
                    .Select(item => new TemplateRequirementValue
                    {
                        Range = CloneRange(item.Range),
                        Operator = item.Operator,
                        Value = item.Value,
                        NegativeValue = item.NegativeValue,
                        PositiveValue = item.PositiveValue,
                        Unit = item.Unit,
                        DisplayText = item.DisplayText,
                        NumberFormat = item.NumberFormat
                    })
                    .ToList(),
                Formula = CloneFormula(source.Formula),
                FormulaVariants = (source.FormulaVariants ?? new List<TemplateFormulaDefinition>())
                    .Where(item => item != null)
                    .Select(CloneFormula)
                    .ToList()
            };
        }

        private static TemplateFormulaDefinition CloneFormula(TemplateFormulaDefinition source)
        {
            if (source == null)
            {
                return null;
            }

            return new TemplateFormulaDefinition
            {
                Formula = source.Formula,
                FormulaR1C1 = source.FormulaR1C1,
                IsConditional = source.IsConditional,
                IsFullyParsed = source.IsFullyParsed,
                References = (source.References ?? new List<TemplateFormulaReference>())
                    .Where(item => item != null)
                    .Select(item => new TemplateFormulaReference
                    {
                        Token = item.Token,
                        SheetName = item.SheetName,
                        Range = CloneRange(item.Range),
                        Role = item.Role
                    })
                    .ToList(),
                Branches = (source.Branches ?? new List<TemplateFormulaBranch>())
                    .Where(item => item != null)
                    .Select(item => new TemplateFormulaBranch
                    {
                        Condition = item.Condition,
                        ValueExpression = item.ValueExpression,
                        DisplayText = item.DisplayText,
                        NumericValue = item.NumericValue,
                        Operator = item.Operator,
                        Unit = item.Unit
                    })
                    .ToList()
            };
        }

        private static CellRange CloneRange(CellRange source)
        {
            if (source == null)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = source.SheetName,
                StartRow = source.StartRow,
                EndRow = source.EndRow,
                StartColumn = source.StartColumn,
                EndColumn = source.EndColumn
            };
        }
    }
}
