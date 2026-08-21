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
            mapping.SectionRange = TrimRangeToDataRows(sheet, mapping.SectionRange);
            mapping.SetpointValueRange = TrimRangeToDataRows(sheet, mapping.SetpointValueRange);
            mapping.MeasurementValueRange = TrimMeasurementRangeToDataRows(sheet, mapping.MeasurementValueRange);
            mapping.SetpointValueRange = AlignStandardValueRangeToMeasurementRows(sheet, mapping.SetpointValueRange, mapping.MeasurementValueRange);
            mapping.StandardValueRange = TrimRangeToDataRows(sheet, mapping.StandardValueRange);
            mapping.StandardValueRange = AlignStandardValueRangeToMeasurementRows(sheet, mapping.StandardValueRange, mapping.MeasurementValueRange);
            mapping.AverageValueRange = TrimRangeToDataRows(sheet, mapping.AverageValueRange);
            mapping.ErrorValueRange = ExpandMergedDataRows(sheet, TrimRangeToDataRows(sheet, mapping.ErrorValueRange));
            mapping.TechnicalRequirementRange = ExpandMergedDataRows(sheet, TrimRangeToDataRows(sheet, mapping.TechnicalRequirementRange));
            mapping.UncertaintyRange = TrimRangeToDataRows(sheet, mapping.UncertaintyRange);
            mapping.RangeValueRange = TrimRangeToDataRows(sheet, mapping.RangeValueRange);
            mapping.ResultRange = TrimRangeToDataRows(sheet, mapping.ResultRange);
            return mapping;
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
