using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropSnapshotProvider
    {
        private SheetSnapshot CaptureSheet(dynamic worksheet)
        {
            dynamic usedRange = worksheet.UsedRange;
            int usedRangeStartRow = Math.Max(1, SafeToInt(usedRange.Row));
            int usedRangeStartColumn = Math.Max(1, SafeToInt(usedRange.Column));
            int rowCount = Math.Max(0, SafeToInt(usedRange.Rows.Count));
            int columnCount = Math.Max(0, SafeToInt(usedRange.Columns.Count));
            var scanAreas = ResolveScanAreas(
                worksheet,
                usedRangeStartRow,
                usedRangeStartColumn,
                rowCount,
                columnCount);
            RecognitionProgress.Report(20, $"正在读取区域 {BuildRangeShape(scanAreas, rowCount, columnCount)}...");

            var sheet = new SheetSnapshot
            {
                Name = SafeToString(worksheet.Name),
                UsedRangeShape = BuildRangeShape(scanAreas, rowCount, columnCount)
            };

            if (rowCount == 0 || columnCount == 0 || scanAreas.Count == 0)
            {
                return sheet;
            }

            var totalCellCount = 0;
            foreach (var area in scanAreas)
            {
                totalCellCount += area.RowCount * area.ColumnCount;
            }

            totalCellCount = Math.Max(1, totalCellCount);
            var processedCellCount = 0;
            var mergeCandidates = new List<CellAddress>();
            foreach (var area in scanAreas)
            {
                dynamic scanRange = worksheet.Range[
                    worksheet.Cells[area.StartRow, area.StartColumn],
                    worksheet.Cells[area.EndRow, area.EndColumn]];
                var values = ReadMatrix(scanRange.Value2, area.RowCount, area.ColumnCount);
                var formulas = ReadMatrix(scanRange.Formula, area.RowCount, area.ColumnCount);
                var formulasR1C1 = ReadMatrix(scanRange.FormulaR1C1, area.RowCount, area.ColumnCount);
                var numberFormats = ReadMatrix(scanRange.NumberFormat, area.RowCount, area.ColumnCount);
                var displayValues = ReadMatrix(scanRange.Text, area.RowCount, area.ColumnCount);

                for (var rowOffset = 1; rowOffset <= area.RowCount; rowOffset++)
                {
                    for (var columnOffset = 1; columnOffset <= area.ColumnCount; columnOffset++)
                    {
                        var row = area.StartRow + rowOffset - 1;
                        var column = area.StartColumn + columnOffset - 1;
                        var rawValueText = SafeToString(values[rowOffset, columnOffset]);
                        var displayText = SafeToString(displayValues[rowOffset, columnOffset]);
                        var text = ResolveCellText(displayText, rawValueText);
                        var formula = NormalizeFormula(formulas[rowOffset, columnOffset]);
                        if (!string.IsNullOrWhiteSpace(rawValueText) ||
                            !string.IsNullOrWhiteSpace(formula))
                        {
                            mergeCandidates.Add(new CellAddress(row, column));
                        }
                        sheet.Cells.Add(new CellMeta
                        {
                            Row = row,
                            Column = column,
                            Text = text,
                            RawValueText = rawValueText,
                            DisplayText = displayText,
                            NumberFormat = SafeToString(numberFormats[rowOffset, columnOffset]),
                            Formula = formula,
                            FormulaR1C1 = NormalizeFormula(formulasR1C1[rowOffset, columnOffset])
                        });

                        processedCellCount++;
                        if (processedCellCount % 800 == 0 || processedCellCount == totalCellCount)
                        {
                            var percent = 20 + (int)Math.Round(processedCellCount * 35d / totalCellCount);
                            RecognitionProgress.Report(percent, $"正在提取单元格 {processedCellCount}/{totalCellCount}...");
                        }
                    }
                }
            }

            RecognitionProgress.Report(58, "正在整理表头与项目区域...");
            List<CellRange> mergedRanges = null;
            if (!CanUsePersistedMergeLayout() ||
                !TryCapturePersistedMergedRanges(worksheet, out mergedRanges))
            {
                mergedRanges = CaptureMergedRangesForCandidates(worksheet, scanAreas, mergeCandidates);
            }
            ApplyMergedRanges(sheet, mergedRanges);
            sheet.Headers.AddRange(BuildHeaderPaths(sheet));
            return sheet;
        }

        private SheetSnapshot CaptureSheetRanges(dynamic worksheet, IReadOnlyList<CellRange> ranges)
        {
            var sheet = new SheetSnapshot
            {
                Name = SafeToString(worksheet.Name),
                UsedRangeShape = string.Join(",", ranges.Select(ToAddress))
            };

            var scanAreas = new List<ScanArea>();
            var mergeCandidates = new List<CellAddress>();
            foreach (var range in ranges)
            {
                var rowCount = range.EndRow - range.StartRow + 1;
                var columnCount = range.EndColumn - range.StartColumn + 1;
                if (rowCount <= 0 || columnCount <= 0)
                {
                    continue;
                }

                scanAreas.Add(new ScanArea
                {
                    StartRow = range.StartRow,
                    StartColumn = range.StartColumn,
                    EndRow = range.EndRow,
                    EndColumn = range.EndColumn
                });

                dynamic excelRange = worksheet.Range[
                    worksheet.Cells[range.StartRow, range.StartColumn],
                    worksheet.Cells[range.EndRow, range.EndColumn]];
                var values = ReadMatrix(excelRange.Value2, rowCount, columnCount);
                var formulas = ReadMatrix(excelRange.Formula, rowCount, columnCount);
                var formulasR1C1 = ReadMatrix(excelRange.FormulaR1C1, rowCount, columnCount);
                var numberFormats = ReadMatrix(excelRange.NumberFormat, rowCount, columnCount);
                var displayValues = ReadMatrix(excelRange.Text, rowCount, columnCount);

                for (var rowOffset = 1; rowOffset <= rowCount; rowOffset++)
                {
                    for (var columnOffset = 1; columnOffset <= columnCount; columnOffset++)
                    {
                        var row = range.StartRow + rowOffset - 1;
                        var column = range.StartColumn + columnOffset - 1;
                        var rawValueText = SafeToString(values[rowOffset, columnOffset]);
                        var displayText = SafeToString(displayValues[rowOffset, columnOffset]);
                        var formula = NormalizeFormula(formulas[rowOffset, columnOffset]);
                        if (!string.IsNullOrWhiteSpace(rawValueText) ||
                            !string.IsNullOrWhiteSpace(formula))
                        {
                            mergeCandidates.Add(new CellAddress(row, column));
                        }
                        sheet.Cells.Add(new CellMeta
                        {
                            Row = row,
                            Column = column,
                            Text = ResolveCellText(displayText, rawValueText),
                            RawValueText = rawValueText,
                            DisplayText = displayText,
                            NumberFormat = SafeToString(numberFormats[rowOffset, columnOffset]),
                            Formula = formula,
                            FormulaR1C1 = NormalizeFormula(formulasR1C1[rowOffset, columnOffset])
                        });
                    }
                }
            }

            List<CellRange> mergedRanges;
            if (!TryCapturePersistedMergedRanges(worksheet, out mergedRanges))
            {
                mergedRanges = CaptureMergedRangesForCandidates(worksheet, scanAreas, mergeCandidates);
            }
            ApplyMergedRanges(sheet, mergedRanges);
            sheet.Headers.AddRange(BuildHeaderPaths(sheet));
            return sheet;
        }

    }
}
