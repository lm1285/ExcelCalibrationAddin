using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed class TemplateFieldDefinitionBuilder
    {
        private static readonly Regex NumberRegex = new Regex(
            @"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?",
            RegexOptions.Compiled);
        private readonly NumberFormatInterpreter _numberFormatInterpreter;

        public TemplateFieldDefinitionBuilder(NumberFormatInterpreter numberFormatInterpreter)
        {
            _numberFormatInterpreter = numberFormatInterpreter ?? new NumberFormatInterpreter();
        }

        public TemplateFieldDefinition Build(
            SheetSnapshot sheet,
            TemplateRegionMapping mapping)
        {
            if (sheet == null || mapping == null)
            {
                return null;
            }

            var ranges = BuildRoleRanges(mapping);
            var definition = new TemplateFieldDefinition
            {
                ProjectName = mapping.ProjectName,
                SectionRange = CloneRange(mapping.SectionRange),
                Headers = BuildFieldHeaders(sheet, mapping, ranges.Values)
            };

            foreach (var pair in ranges)
            {
                if (pair.Value == null) continue;
                definition.Regions.Add(BuildRegion(sheet, mapping, pair.Key, pair.Value, ranges));
            }

            return definition;
        }

        public TemplateFieldDefinition Build(SheetSnapshot sheet, MeasurementRule rule)
        {
            if (sheet == null || rule?.TargetRange == null)
            {
                return null;
            }

            var ranges = new[]
            {
                rule.TargetRange,
                rule.SetpointSource?.Range,
                rule.StandardValueSource?.Range,
                rule.AverageSource?.Range,
                rule.ErrorSource?.Range,
                rule.MpeSource?.Range,
                rule.RangeSource?.Range,
                rule.UncertaintySource?.Range,
                rule.ResultSource?.Range
            }.Where(range => range != null).ToList();
            var existingSection = rule.TemplateDefinition?.SectionRange;
            var section = existingSection != null
                ? CloneRange(existingSection)
                : new CellRange
                {
                    SheetName = rule.TargetRange.SheetName,
                    StartRow = Math.Max(1, ranges.Min(range => range.StartRow) - 4),
                    EndRow = ranges.Max(range => range.EndRow),
                    StartColumn = ranges.Min(range => range.StartColumn),
                    EndColumn = ranges.Max(range => range.EndColumn)
                };
            return Build(sheet, new TemplateRegionMapping
            {
                ProjectName = string.IsNullOrWhiteSpace(rule.FieldAlias) ? rule.FieldName : rule.FieldAlias,
                SectionRange = section,
                SetpointValueRange = CloneRange(rule.SetpointSource?.Range) ??
                    CombineRowRanges(rule.RowMappings, mapping => mapping.SetpointValueRange),
                StandardValueRange = CloneRange(rule.StandardValueSource?.Range) ??
                    CombineRowRanges(rule.RowMappings, mapping => mapping.StandardValueRange),
                MeasurementValueRange = CloneRange(rule.TargetRange),
                AverageValueRange = CloneRange(rule.AverageSource?.Range) ??
                    CombineRowRanges(rule.RowMappings, mapping => mapping.AverageRange),
                ErrorValueRange = CloneRange(rule.ErrorSource?.Range) ??
                    CombineRowRanges(rule.RowMappings, mapping => mapping.ErrorRange),
                TechnicalRequirementRange = CloneRange(rule.MpeSource?.Range) ??
                    CombineRowRanges(rule.RowMappings, mapping => mapping.TechnicalRequirementRange),
                RangeValueRange = CloneRange(rule.RangeSource?.Range) ??
                    CombineRowRanges(rule.RowMappings, mapping => mapping.RangeValueRange),
                UncertaintyRange = CloneRange(rule.UncertaintySource?.Range) ??
                    CombineRowRanges(rule.RowMappings, mapping => mapping.UncertaintyRange),
                ResultRange = CloneRange(rule.ResultSource?.Range) ??
                    CombineRowRanges(rule.RowMappings, mapping => mapping.ResultRange)
            });
        }

        private static CellRange CombineRowRanges(
            IEnumerable<MeasurementRowMapping> mappings,
            Func<MeasurementRowMapping, CellRange> selector)
        {
            var ranges = (mappings ?? Enumerable.Empty<MeasurementRowMapping>())
                .Where(mapping => mapping != null)
                .Select(selector)
                .Where(range => range != null)
                .ToList();
            if (ranges.Count == 0) return null;
            var sheetName = ranges[0].SheetName;
            if (ranges.Any(range => !string.Equals(range.SheetName, sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                return null;
            }

            return new CellRange
            {
                SheetName = sheetName,
                StartRow = ranges.Min(range => range.StartRow),
                EndRow = ranges.Max(range => range.EndRow),
                StartColumn = ranges.Min(range => range.StartColumn),
                EndColumn = ranges.Max(range => range.EndColumn)
            };
        }

        private TemplateRegionDefinition BuildRegion(
            SheetSnapshot sheet,
            TemplateRegionMapping mapping,
            TemplateRegionRole role,
            CellRange range,
            IReadOnlyDictionary<TemplateRegionRole, CellRange> roleRanges)
        {
            var headerRange = ResolveHeaderRange(mapping.SectionRange, range);
            var headerCells = GetLogicalCells(sheet, headerRange)
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Anchor?.Text))
                .OrderBy(cell => cell.Range.StartRow)
                .ThenBy(cell => cell.Range.StartColumn)
                .ToList();
            var contentCells = GetLogicalCells(sheet, range)
                .OrderBy(cell => cell.Range.StartRow)
                .ThenBy(cell => cell.Range.StartColumn)
                .ToList();
            var formatCell = contentCells
                .Select(cell => cell.Anchor)
                .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell?.NumberFormat));
            var formulaCells = contentCells
                .Select(cell => cell.Anchor)
                .Where(cell => !string.IsNullOrWhiteSpace(cell?.Formula))
                .GroupBy(cell => string.IsNullOrWhiteSpace(cell.FormulaR1C1) ? cell.Formula : cell.FormulaR1C1,
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();

            var displayUnit = TemplateUnitParser.Extract(contentCells.SelectMany(cell => new[]
            {
                cell.Anchor?.DisplayText,
                cell.Anchor?.Text,
                cell.Anchor?.RawValueText
            }));
            var formatUnit = _numberFormatInterpreter.Interpret(formatCell?.NumberFormat).UnitSuffix;
            if (string.IsNullOrWhiteSpace(formatUnit))
            {
                formatUnit = TemplateUnitParser.Extract(formatCell?.NumberFormat);
            }
            var headerUnit = TemplateUnitParser.Extract(headerCells.Select(cell => cell.Anchor?.Text));
            var unit = !string.IsNullOrWhiteSpace(displayUnit)
                ? displayUnit
                : !string.IsNullOrWhiteSpace(formatUnit)
                    ? formatUnit
                    : headerUnit;
            var unitSource = !string.IsNullOrWhiteSpace(displayUnit)
                ? TemplateUnitSource.CellDisplay
                : !string.IsNullOrWhiteSpace(formatUnit)
                    ? TemplateUnitSource.CellFormat
                    : !string.IsNullOrWhiteSpace(headerUnit)
                        ? TemplateUnitSource.RegionHeader
                        : TemplateUnitSource.None;

            var definition = new TemplateRegionDefinition
            {
                Role = role,
                Range = CloneRange(range),
                HeaderRange = CloneRange(headerRange),
                HeaderPath = headerCells
                    .Select(cell => (cell.Anchor?.Text ?? string.Empty).Trim())
                    .Where(text => !string.IsNullOrWhiteSpace(text))
                    .Distinct()
                    .ToList(),
                Unit = unit,
                UnitSource = unitSource,
                NumberFormat = formatCell?.NumberFormat ?? string.Empty,
                FormulaVariants = formulaCells
                    .Select(cell => TemplateFormulaParser.Parse(sheet, cell, roleRanges))
                    .Where(formula => formula != null)
                    .ToList()
            };
            definition.Formula = definition.FormulaVariants.FirstOrDefault();

            if (definition.Formula?.Branches?.Any(branch => !string.IsNullOrWhiteSpace(branch.Unit)) == true &&
                string.IsNullOrWhiteSpace(definition.Unit))
            {
                definition.Unit = definition.Formula.Branches
                    .Select(branch => branch.Unit)
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
                definition.UnitSource = TemplateUnitSource.FormulaBranch;
            }
            definition.Units = new[] { displayUnit, formatUnit, headerUnit, definition.Unit }
                .Concat(definition.FormulaVariants.SelectMany(formula =>
                    formula.Branches.Select(branch => branch.Unit)))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (role == TemplateRegionRole.TechnicalRequirement)
            {
                ResolveSplitRequirementRanges(contentCells, range, out var operatorRange, out var valueRange);
                definition.OperatorRange = operatorRange;
                definition.ValueRange = valueRange;
                definition.RequirementValues = BuildRequirementValues(
                    contentCells,
                    headerCells,
                    operatorRange);
            }

            return definition;
        }

        private static List<TemplateRequirementValue> BuildRequirementValues(
            IReadOnlyList<LogicalCellRange> contentCells,
            IReadOnlyList<LogicalCellRange> headerCells,
            CellRange operatorRange)
        {
            var sharedOperator = contentCells
                .Where(cell => RangeContains(operatorRange, cell.Range.StartRow, cell.Range.StartColumn))
                .Select(cell => RequirementTextParser.Parse(cell.Anchor).Operator)
                .FirstOrDefault(value => value != TechnicalRequirementOperator.None);
            var headerUnit = TemplateUnitParser.Extract(headerCells.Select(cell => cell.Anchor?.Text));
            var result = new List<TemplateRequirementValue>();
            foreach (var logicalCell in contentCells)
            {
                var anchor = logicalCell.Anchor;
                var parsed = RequirementTextParser.Parse(anchor);
                var displayText = string.Join(" ", new[]
                {
                    anchor?.DisplayText,
                    anchor?.Text,
                    anchor?.RawValueText
                }.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct());
                var matches = NumberRegex.Matches(displayText)
                    .Cast<Match>()
                    .Select(match =>
                    {
                        double value;
                        return double.TryParse(
                            match.Value,
                            System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out value)
                            ? (double?)value
                            : null;
                    })
                    .Where(value => value.HasValue)
                    .Select(value => value.Value)
                    .ToList();
                if (matches.Count == 0)
                {
                    continue;
                }

                var negative = matches.FirstOrDefault(value => value < 0);
                var positive = matches.FirstOrDefault(value => value > 0);
                result.Add(new TemplateRequirementValue
                {
                    Range = CloneRange(logicalCell.Range),
                    Operator = parsed.Operator != TechnicalRequirementOperator.None
                        ? parsed.Operator
                        : sharedOperator,
                    Value = Math.Abs(matches[0]),
                    NegativeValue = matches.Any(value => value < 0) ? (double?)Math.Abs(negative) : null,
                    PositiveValue = matches.Count > 1 && matches.Any(value => value > 0)
                        ? (double?)Math.Abs(positive)
                        : null,
                    Unit = TemplateUnitParser.Extract(
                        anchor?.DisplayText,
                        anchor?.Text,
                        anchor?.NumberFormat,
                        headerUnit),
                    DisplayText = displayText,
                    NumberFormat = anchor?.NumberFormat ?? string.Empty
                });
            }

            return result;
        }

        private static List<TemplateHeaderDefinition> BuildFieldHeaders(
            SheetSnapshot sheet,
            TemplateRegionMapping mapping,
            IEnumerable<CellRange> ranges)
        {
            var validRanges = (ranges ?? Enumerable.Empty<CellRange>())
                .Where(range => range != null)
                .ToList();
            if (mapping.SectionRange == null || validRanges.Count == 0)
            {
                return new List<TemplateHeaderDefinition>();
            }

            var endRow = validRanges.Min(range => range.StartRow) - 1;
            if (endRow < mapping.SectionRange.StartRow)
            {
                return new List<TemplateHeaderDefinition>();
            }

            var headerRange = new CellRange
            {
                SheetName = mapping.SectionRange.SheetName,
                StartRow = mapping.SectionRange.StartRow,
                EndRow = endRow,
                StartColumn = mapping.SectionRange.StartColumn,
                EndColumn = mapping.SectionRange.EndColumn
            };
            return GetLogicalCells(sheet, headerRange)
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Anchor?.Text))
                .GroupBy(cell => RangeKey(cell.Range))
                .Select(group => group.First())
                .OrderBy(cell => cell.Range.StartRow)
                .ThenBy(cell => cell.Range.StartColumn)
                .Select(cell => new TemplateHeaderDefinition
                {
                    Level = cell.Range.StartRow - headerRange.StartRow + 1,
                    Range = CloneRange(cell.Range),
                    Text = (cell.Anchor?.Text ?? string.Empty).Trim(),
                    Unit = TemplateUnitParser.Extract(cell.Anchor?.Text, cell.Anchor?.NumberFormat),
                    NumberFormat = cell.Anchor?.NumberFormat ?? string.Empty
                })
                .ToList();
        }

        private static void ResolveSplitRequirementRanges(
            IReadOnlyList<LogicalCellRange> cells,
            CellRange source,
            out CellRange operatorRange,
            out CellRange valueRange)
        {
            operatorRange = null;
            valueRange = null;
            if (source == null || source.EndColumn <= source.StartColumn)
            {
                return;
            }

            var operatorColumns = new List<int>();
            var valueColumns = new List<int>();
            foreach (var columnGroup in cells.GroupBy(cell => cell.Range.StartColumn))
            {
                var hasOperator = columnGroup.Any(cell =>
                    RequirementTextParser.Parse(cell.Anchor).Operator != TechnicalRequirementOperator.None);
                var hasNumber = columnGroup.Any(cell => NumberRegex.IsMatch(string.Join(" ", new[]
                {
                    cell.Anchor?.DisplayText,
                    cell.Anchor?.Text,
                    cell.Anchor?.RawValueText
                }.Where(text => !string.IsNullOrWhiteSpace(text)))));
                if (hasOperator && !hasNumber) operatorColumns.Add(columnGroup.Key);
                if (hasNumber) valueColumns.Add(columnGroup.Key);
            }

            if (operatorColumns.Count == 0 || valueColumns.Count == 0 ||
                operatorColumns.Min() >= valueColumns.Max())
            {
                return;
            }

            operatorRange = ColumnRange(source, operatorColumns.Min(), operatorColumns.Max());
            valueRange = ColumnRange(source, valueColumns.Min(), valueColumns.Max());
        }

        private static CellRange ResolveHeaderRange(CellRange section, CellRange range)
        {
            if (range == null) return null;
            var startRow = Math.Max(section?.StartRow ?? 1, range.StartRow - 4);
            var endRow = range.StartRow - 1;
            return endRow < startRow
                ? null
                : new CellRange
                {
                    SheetName = range.SheetName,
                    StartRow = startRow,
                    EndRow = endRow,
                    StartColumn = range.StartColumn,
                    EndColumn = range.EndColumn
                };
        }

        private static Dictionary<TemplateRegionRole, CellRange> BuildRoleRanges(TemplateRegionMapping mapping)
        {
            return new Dictionary<TemplateRegionRole, CellRange>
            {
                [TemplateRegionRole.SetpointValue] = mapping.SetpointValueRange,
                [TemplateRegionRole.StandardValue] = mapping.StandardValueRange,
                [TemplateRegionRole.MeasurementValue] = mapping.MeasurementValueRange,
                [TemplateRegionRole.AverageValue] = mapping.AverageValueRange,
                [TemplateRegionRole.ErrorValue] = mapping.ErrorValueRange,
                [TemplateRegionRole.TechnicalRequirement] = mapping.TechnicalRequirementRange,
                [TemplateRegionRole.RangeValue] = mapping.RangeValueRange,
                [TemplateRegionRole.Uncertainty] = mapping.UncertaintyRange,
                [TemplateRegionRole.Result] = mapping.ResultRange
            };
        }

        private static IReadOnlyList<LogicalCellRange> GetLogicalCells(SheetSnapshot sheet, CellRange range)
        {
            return range == null
                ? new List<LogicalCellRange>()
                : MergedCellLogicalRangeResolver.GetContentCells(sheet, range);
        }

        private static CellRange ColumnRange(CellRange source, int startColumn, int endColumn)
        {
            return new CellRange
            {
                SheetName = source.SheetName,
                StartRow = source.StartRow,
                EndRow = source.EndRow,
                StartColumn = startColumn,
                EndColumn = endColumn
            };
        }

        private static string RangeKey(CellRange range)
        {
            return range == null
                ? string.Empty
                : $"{range.StartRow}:{range.StartColumn}:{range.EndRow}:{range.EndColumn}";
        }

        private static bool RangeContains(CellRange range, int row, int column)
        {
            return range != null &&
                row >= range.StartRow && row <= range.EndRow &&
                column >= range.StartColumn && column <= range.EndColumn;
        }

        private static CellRange CloneRange(CellRange range)
        {
            return range == null
                ? null
                : new CellRange
                {
                    SheetName = range.SheetName,
                    StartRow = range.StartRow,
                    EndRow = range.EndRow,
                    StartColumn = range.StartColumn,
                    EndColumn = range.EndColumn
                };
        }
    }
}
