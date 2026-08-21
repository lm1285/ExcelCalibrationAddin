using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropWriter
    {
        private static bool TryWriteDenseRange(dynamic worksheet, Contracts.CellRange range, IReadOnlyList<string> values)
        {
            var rowCount = range.EndRow - range.StartRow + 1;
            var columnCount = range.EndColumn - range.StartColumn + 1;
            if (rowCount <= 0 || columnCount <= 0 || values.Count != rowCount * columnCount)
            {
                return false;
            }

            try
            {
                dynamic target = worksheet.Range[
                    worksheet.Cells[range.StartRow, range.StartColumn],
                    worksheet.Cells[range.EndRow, range.EndColumn]];
                if (RangeHasMergedCells(target))
                {
                    return false;
                }

                var matrix = new object[rowCount, columnCount];
                var valueIndex = 0;
                for (var row = 0; row < rowCount; row++)
                {
                    for (var column = 0; column < columnCount; column++)
                    {
                        matrix[row, column] = values[valueIndex++];
                    }
                }

                target.Value2 = matrix;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool RangeHasMergedCells(dynamic range)
        {
            try
            {
                var mergeState = range.MergeCells;
                if (mergeState == null)
                {
                    return true;
                }

                return Convert.ToBoolean(mergeState);
            }
            catch
            {
                return true;
            }
        }

    }
}
