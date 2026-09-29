using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Recognition
{
    public static class TemplateRegionMappingNormalizer
    {
        public static TemplateRegionMapping Normalize(SheetSnapshot sheet, TemplateRegionMapping mapping)
        {
            // Keep the recognized section bounds before trimming blank/note rows.
            // The trimmed section is useful for matching, but must not allow a
            // merged field to expand into the next calibration item.
            var sectionStartRow = mapping.SectionRange?.StartRow;
            var sectionEndRow = mapping.SectionRange?.EndRow;
            mapping.SectionRange = TrimRangeToDataRows(sheet, mapping.SectionRange);
            mapping.SetpointValueRange = TrimRangeToDataRows(sheet, mapping.SetpointValueRange);
            mapping.MeasurementValueRange = TrimMeasurementRangeToDataRows(sheet, mapping.MeasurementValueRange);
            mapping.SetpointValueRange = AlignStandardValueRangeToMeasurementRows(sheet, mapping.SetpointValueRange, mapping.MeasurementValueRange);
            mapping.StandardValueRange = TrimRangeToDataRows(sheet, mapping.StandardValueRange);
            mapping.StandardValueRange = AlignStandardValueRangeToMeasurementRows(sheet, mapping.StandardValueRange, mapping.MeasurementValueRange);
            mapping.AverageValueRange = ExpandMergedDataColumns(sheet, TrimRangeToDataRows(sheet, mapping.AverageValueRange));
            mapping.ErrorValueRange = ExpandMergedDataColumns(sheet, ExpandMergedDataRows(sheet, TrimRangeToDataRows(sheet, mapping.ErrorValueRange)));
            mapping.TechnicalRequirementRange = ExpandMergedDataColumns(sheet, ExpandMergedDataRows(sheet, TrimRangeToDataRows(sheet, mapping.TechnicalRequirementRange)));
            mapping.UncertaintyRange = TrimRangeToDataRows(sheet, mapping.UncertaintyRange);
            mapping.RangeValueRange = TrimRangeToDataRows(sheet, mapping.RangeValueRange);
            mapping.ResultRange = TrimRangeToDataRows(sheet, mapping.ResultRange);
            mapping.SetpointValueRange = ClampToSection(mapping.SetpointValueRange, sectionStartRow, sectionEndRow);
            mapping.StandardValueRange = ClampToSection(mapping.StandardValueRange, sectionStartRow, sectionEndRow);
            mapping.MeasurementValueRange = ClampToSection(mapping.MeasurementValueRange, sectionStartRow, sectionEndRow);
            mapping.AverageValueRange = ClampToSection(mapping.AverageValueRange, sectionStartRow, sectionEndRow);
            mapping.ErrorValueRange = ClampToSection(mapping.ErrorValueRange, sectionStartRow, sectionEndRow);
            mapping.TechnicalRequirementRange = ClampToSection(mapping.TechnicalRequirementRange, sectionStartRow, sectionEndRow);
            mapping.UncertaintyRange = ClampToSection(mapping.UncertaintyRange, sectionStartRow, sectionEndRow);
            mapping.ResultRange = ClampToSection(mapping.ResultRange, sectionStartRow, sectionEndRow);
            return mapping;
        }

        private static CellRange ClampToSection(CellRange range, int? sectionStartRow, int? sectionEndRow)
        {
            if (range == null || !sectionStartRow.HasValue || !sectionEndRow.HasValue)
            {
                return range;
            }

            var startRow = Math.Max(range.StartRow, sectionStartRow.Value);
            var endRow = Math.Min(range.EndRow, sectionEndRow.Value);
            if (endRow < startRow)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = startRow,
                EndRow = endRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }

        private static CellRange TrimRangeToDataRows(SheetSnapshot sheet, CellRange range)
        {
            if (range == null)
            {
                return null;
            }

            var currentEndRow = Math.Max(range.StartRow, range.EndRow);
            while (currentEndRow > range.StartRow &&
                   (SheetRowContentAnalyzer.IsTrailingNoteRow(sheet, currentEndRow) ||
                    !SheetRowContentAnalyzer.HasDataInRangeRow(sheet, currentEndRow, range.StartColumn, range.EndColumn)))
            {
                currentEndRow--;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                EndRow = currentEndRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }

        private static CellRange ExpandMergedDataRows(SheetSnapshot sheet, CellRange range)
        {
            if (sheet == null || range == null)
            {
                return range;
            }

            var endRow = range.EndRow;
            var changed = true;
            while (changed)
            {
                changed = false;
                foreach (var cell in sheet.Cells.Where(item =>
                    item.Column >= range.StartColumn &&
                    item.Column <= range.EndColumn &&
                    item.MergeRange != null &&
                    item.MergeRange.StartRow <= endRow &&
                    item.MergeRange.EndRow > endRow))
                {
                    endRow = cell.MergeRange.EndRow;
                    changed = true;
                }
            }

            if (endRow == range.EndRow)
            {
                return range;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                EndRow = endRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }

        private static CellRange ExpandMergedDataColumns(SheetSnapshot sheet, CellRange range)
        {
            if (sheet == null || range == null)
            {
                return range;
            }

            var startColumn = range.StartColumn;
            var endColumn = range.EndColumn;
            var changed = true;
            while (changed)
            {
                changed = false;
                foreach (var cell in sheet.Cells.Where(item =>
                    item.Row >= range.StartRow &&
                    item.Row <= range.EndRow &&
                    item.MergeRange != null &&
                    item.MergeRange.StartColumn <= endColumn &&
                    item.MergeRange.EndColumn >= startColumn))
                {
                    if (cell.MergeRange.StartColumn < startColumn)
                    {
                        startColumn = cell.MergeRange.StartColumn;
                        changed = true;
                    }
                    if (cell.MergeRange.EndColumn > endColumn)
                    {
                        endColumn = cell.MergeRange.EndColumn;
                        changed = true;
                    }
                }
            }

            if (startColumn == range.StartColumn && endColumn == range.EndColumn)
            {
                return range;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                EndRow = range.EndRow,
                StartColumn = startColumn,
                EndColumn = endColumn
            };
        }

        private static CellRange AlignStandardValueRangeToMeasurementRows(
            SheetSnapshot sheet,
            CellRange standardRange,
            CellRange measurementRange)
        {
            if (sheet == null ||
                standardRange == null ||
                measurementRange == null ||
                measurementRange.EndRow <= measurementRange.StartRow)
            {
                return standardRange;
            }

            var standardRows = new List<int>();
            for (var row = measurementRange.StartRow; row <= measurementRange.EndRow; row++)
            {
                if (SheetRowContentAnalyzer.HasNumericDataInRangeRow(sheet, row, standardRange.StartColumn, standardRange.EndColumn))
                {
                    standardRows.Add(row);
                }
            }

            if (standardRows.Count <= 1)
            {
                return standardRange;
            }

            return new CellRange
            {
                SheetName = standardRange.SheetName,
                StartRow = standardRows.Min(),
                EndRow = standardRows.Max(),
                StartColumn = standardRange.StartColumn,
                EndColumn = standardRange.EndColumn
            };
        }

        private static CellRange TrimMeasurementRangeToDataRows(SheetSnapshot sheet, CellRange range)
        {
            if (range == null)
            {
                return null;
            }

            var currentEndRow = Math.Max(range.StartRow, range.EndRow);
            while (currentEndRow > range.StartRow &&
                   (SheetRowContentAnalyzer.IsTrailingNoteRow(sheet, currentEndRow) ||
                    (!SheetRowContentAnalyzer.HasDataInRangeRow(sheet, currentEndRow, range.StartColumn, range.EndColumn) &&
                     SheetRowContentAnalyzer.CountWritableTemplateCellsInRow(sheet, currentEndRow, range.StartColumn, range.EndColumn) <= 0)))
            {
                currentEndRow--;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                EndRow = currentEndRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }
    }
}
