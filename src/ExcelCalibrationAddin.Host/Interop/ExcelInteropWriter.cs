using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropWriter : IWorkbookWriter
    {
        private readonly dynamic _workbook;

        public ExcelInteropWriter(dynamic workbook)
        {
            _workbook = workbook;
        }

        public void Write(Contracts.CellRange range, IReadOnlyList<string> values)
        {
            Write(range, null, values);
        }

        public void Write(Contracts.CellRange range, IReadOnlyList<Contracts.CellAddress> writableCells, IReadOnlyList<string> values)
        {
            ValidateWriteRequest(range, values);

            dynamic worksheet = _workbook.Worksheets[range.SheetName];
            if (TryWriteKnownCells(worksheet, range, writableCells, values))
            {
                return;
            }

            if (TryWriteDenseRange(worksheet, range, values))
            {
                return;
            }

            EnsureWritableRange(worksheet, range);

            var valueIndex = 0;

            for (var row = range.StartRow; row <= range.EndRow; row++)
            {
                for (var column = range.StartColumn; column <= range.EndColumn; column++)
                {
                    if (valueIndex >= values.Count)
                    {
                        return;
                    }

                    dynamic cell = worksheet.Cells[row, column];
                    if (IsTopLeftOfMergeArea(cell, row, column))
                    {
                        EnsureWritableCell(cell, row, column);
                        cell.Value2 = values[valueIndex++];
                    }
                    else if (!IsMerged(cell))
                    {
                        EnsureWritableCell(cell, row, column);
                        cell.Value2 = values[valueIndex++];
                    }
                }
            }

            if (valueIndex != values.Count)
            {
                throw new InvalidOperationException($"写入区域可写单元格数量不足。已写入 {valueIndex} 个，待写入 {values.Count} 个。");
            }
        }

    }
}
