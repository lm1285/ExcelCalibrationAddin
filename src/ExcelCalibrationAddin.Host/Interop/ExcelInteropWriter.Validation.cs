using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropWriter
    {
        private static void ValidateWriteRequest(Contracts.CellRange range, IReadOnlyList<string> values)
        {
            if (range == null)
            {
                throw new InvalidOperationException("写入区域为空。");
            }

            if (string.IsNullOrWhiteSpace(range.SheetName))
            {
                throw new InvalidOperationException("写入区域缺少工作表名称。");
            }

            if (range.StartRow <= 0 ||
                range.StartColumn <= 0 ||
                range.EndRow < range.StartRow ||
                range.EndColumn < range.StartColumn)
            {
                throw new InvalidOperationException($"写入区域无效：{range}");
            }

            if (values == null || values.Count == 0)
            {
                throw new InvalidOperationException("没有可写入的随机数。");
            }

            if (values.Any(string.IsNullOrWhiteSpace))
            {
                throw new InvalidOperationException("随机数列表包含空值，写入已取消。");
            }
        }

        private static bool IsMerged(dynamic cell)
        {
            try
            {
                return cell.MergeCells;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsTopLeftOfMergeArea(dynamic cell, int row, int column)
        {
            try
            {
                if (!cell.MergeCells)
                {
                    return true;
                }

                dynamic mergeArea = cell.MergeArea;
                return mergeArea.Row == row && mergeArea.Column == column;
            }
            catch
            {
                return true;
            }
        }

        private static void EnsureWritableCells(dynamic worksheet, IReadOnlyList<Contracts.CellAddress> cells)
        {
            foreach (var cell in cells)
            {
                EnsureWritableCell(worksheet.Cells[cell.Row, cell.Column], cell.Row, cell.Column);
            }
        }

        private static void EnsureWritableRange(dynamic worksheet, Contracts.CellRange range)
        {
            // Random generation intentionally replaces every configured measurement cell, including formulas.
        }

        private static void EnsureWritableCell(dynamic cell, int row, int column)
        {
            // Random generation intentionally replaces every configured measurement cell, including formulas.
        }

    }
}
