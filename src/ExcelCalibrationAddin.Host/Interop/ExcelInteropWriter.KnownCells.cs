using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropWriter
    {
        private static bool TryWriteKnownCells(
            dynamic worksheet,
            Contracts.CellRange range,
            IReadOnlyList<Contracts.CellAddress> writableCells,
            IReadOnlyList<string> values)
        {
            var cells = (writableCells ?? new List<Contracts.CellAddress>())
                .Where(cell => cell != null && cell.Row > 0 && cell.Column > 0)
                .Take(values.Count)
                .ToList();
            if (cells.Count != values.Count)
            {
                return false;
            }

            EnsureWritableCells(worksheet, cells);

            if (TryWriteKnownDenseRange(worksheet, range, cells, values))
            {
                return true;
            }

            if (TryWriteKnownBoundingRange(worksheet, cells, values))
            {
                return true;
            }

            if (TryWriteKnownRuns(worksheet, cells, values))
            {
                return true;
            }

            for (var index = 0; index < values.Count; index++)
            {
                var cell = cells[index];
                dynamic targetCell = worksheet.Cells[cell.Row, cell.Column];
                EnsureWritableCell(targetCell, cell.Row, cell.Column);
                targetCell.Value2 = values[index];
            }

            return true;
        }

        private static bool TryWriteKnownBoundingRange(
            dynamic worksheet,
            IReadOnlyList<Contracts.CellAddress> cells,
            IReadOnlyList<string> values)
        {
            if (cells.Count == 0 || cells.Count != values.Count)
            {
                return false;
            }

            try
            {
                var startRow = cells.Min(cell => cell.Row);
                var endRow = cells.Max(cell => cell.Row);
                var startColumn = cells.Min(cell => cell.Column);
                var endColumn = cells.Max(cell => cell.Column);
                var rowCount = endRow - startRow + 1;
                var columnCount = endColumn - startColumn + 1;
                dynamic target = worksheet.Range[
                    worksheet.Cells[startRow, startColumn],
                    worksheet.Cells[endRow, endColumn]];
                if (RangeHasMergedCells(target))
                {
                    return false;
                }

                var matrix = ToMatrix(target.Formula, rowCount, columnCount);
                for (var index = 0; index < cells.Count; index++)
                {
                    matrix[cells[index].Row - startRow, cells[index].Column - startColumn] = values[index];
                }

                target.Formula = matrix;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static object[,] ToMatrix(object raw, int rowCount, int columnCount)
        {
            var matrix = new object[rowCount, columnCount];
            if (raw is object[,] values)
            {
                for (var row = 0; row < rowCount; row++)
                {
                    for (var column = 0; column < columnCount; column++)
                    {
                        matrix[row, column] = values[row + 1, column + 1];
                    }
                }

                return matrix;
            }

            matrix[0, 0] = raw;
            return matrix;
        }

        private static bool TryWriteKnownDenseRange(
            dynamic worksheet,
            Contracts.CellRange range,
            IReadOnlyList<Contracts.CellAddress> writableCells,
            IReadOnlyList<string> values)
        {
            var rowCount = range.EndRow - range.StartRow + 1;
            var columnCount = range.EndColumn - range.StartColumn + 1;
            if (rowCount <= 0 || columnCount <= 0 || values.Count != rowCount * columnCount)
            {
                return false;
            }

            for (var row = range.StartRow; row <= range.EndRow; row++)
            {
                for (var column = range.StartColumn; column <= range.EndColumn; column++)
                {
                    var offset = ((row - range.StartRow) * columnCount) + (column - range.StartColumn);
                    if (writableCells[offset].Row != row || writableCells[offset].Column != column)
                    {
                        return false;
                    }
                }
            }

            return TryWriteDenseRange(worksheet, range, values);
        }

        private static bool TryWriteKnownRuns(
            dynamic worksheet,
            IReadOnlyList<Contracts.CellAddress> writableCells,
            IReadOnlyList<string> values)
        {
            if (writableCells.Count != values.Count || values.Count <= 1)
            {
                return false;
            }

            var index = 0;
            while (index < values.Count)
            {
                var start = writableCells[index];
                var end = start;
                var runStart = index;
                var runLength = 1;

                while (index + runLength < values.Count)
                {
                    var current = writableCells[index + runLength];
                    var previous = writableCells[index + runLength - 1];
                    var sameRowNextColumn = current.Row == previous.Row && current.Column == previous.Column + 1;
                    var sameColumnNextRow = current.Column == previous.Column && current.Row == previous.Row + 1;
                    if (!sameRowNextColumn && !sameColumnNextRow)
                    {
                        break;
                    }

                    if (runLength == 1)
                    {
                        end = current;
                    }
                    else if (!IsSameDirection(start, end, previous, current))
                    {
                        break;
                    }

                    end = current;
                    runLength++;
                }

                if (runLength == 1)
                {
                    dynamic cell = worksheet.Cells[start.Row, start.Column];
                    EnsureWritableCell(cell, start.Row, start.Column);
                    cell.Value2 = values[runStart];
                }
                else
                {
                    WriteRun(worksheet, start, end, values, runStart, runLength);
                }

                index += runLength;
            }

            return true;
        }

        private static bool IsSameDirection(
            Contracts.CellAddress start,
            Contracts.CellAddress end,
            Contracts.CellAddress previous,
            Contracts.CellAddress current)
        {
            var rowRun = start.Row == end.Row && current.Row == previous.Row && current.Column == previous.Column + 1;
            var columnRun = start.Column == end.Column && current.Column == previous.Column && current.Row == previous.Row + 1;
            return rowRun || columnRun;
        }

        private static void WriteRun(
            dynamic worksheet,
            Contracts.CellAddress start,
            Contracts.CellAddress end,
            IReadOnlyList<string> values,
            int valueStart,
            int valueCount)
        {
            dynamic target = worksheet.Range[
                worksheet.Cells[start.Row, start.Column],
                worksheet.Cells[end.Row, end.Column]];
            if (start.Row == end.Row)
            {
                var matrix = new object[1, valueCount];
                for (var index = 0; index < valueCount; index++)
                {
                    matrix[0, index] = values[valueStart + index];
                }

                target.Value2 = matrix;
                return;
            }

            var columnMatrix = new object[valueCount, 1];
            for (var index = 0; index < valueCount; index++)
            {
                columnMatrix[index, 0] = values[valueStart + index];
            }

            target.Value2 = columnMatrix;
        }

    }
}
