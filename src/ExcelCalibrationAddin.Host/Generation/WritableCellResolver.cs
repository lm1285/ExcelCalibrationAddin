using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Generation
{
    public static class WritableCellResolver
    {
        public static WritableCellResolution Resolve(WorkbookSnapshot snapshot, CellRange range)
        {
            var result = new List<CellAddress>();
            if (range == null ||
                range.StartRow <= 0 ||
                range.StartColumn <= 0 ||
                range.EndRow < range.StartRow ||
                range.EndColumn < range.StartColumn)
            {
                return new WritableCellResolution
                {
                    Cells = result,
                    FailureReason = BuildInvalidTargetRangeReason(range)
                };
            }

            var sheet = snapshot?.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, range.SheetName, StringComparison.OrdinalIgnoreCase));
            if (snapshot == null)
            {
                return new WritableCellResolution
                {
                    Cells = result,
                    FailureReason = "未能读取当前工作簿快照，无法确认测量值区域是否可写。"
                };
            }

            if (sheet == null)
            {
                return new WritableCellResolution
                {
                    Cells = result,
                    FailureReason = $"当前工作簿快照中未找到工作表“{range.SheetName}”。请确认模板保存的工作表名与当前工作簿一致。"
                };
            }

            var cellLookup = (sheet.Cells ?? new List<CellMeta>())
                .Where(cell => cell != null)
                .GroupBy(cell => BuildCellKey(cell.Row, cell.Column))
                .ToDictionary(group => group.Key, group => group.Last());
            var totalCells = CountRangeCells(range);
            var missingCells = 0;
            var mergedCoveredCells = 0;

            for (var row = range.StartRow; row <= range.EndRow; row++)
            {
                for (var column = range.StartColumn; column <= range.EndColumn; column++)
                {
                    CellMeta cell = null;
                    cellLookup.TryGetValue(BuildCellKey(row, column), out cell);
                    if (cell == null)
                    {
                        missingCells++;
                    }

                    if (cell?.MergeRange != null &&
                        (cell.MergeRange.StartRow != row || cell.MergeRange.StartColumn != column))
                    {
                        mergedCoveredCells++;
                        continue;
                    }

                    result.Add(new CellAddress { Row = row, Column = column });
                }
            }

            return new WritableCellResolution
            {
                Cells = result,
                FailureReason = result.Count > 0
                    ? null
                    : BuildNoWritableCellReason(range, totalCells, missingCells, mergedCoveredCells)
            };
        }

        public static int CountRangeCells(CellRange range)
        {
            if (!GenerationRuleValidator.HasValidRange(range))
            {
                return 0;
            }

            return (range.EndRow - range.StartRow + 1) * (range.EndColumn - range.StartColumn + 1);
        }

        private static string BuildInvalidTargetRangeReason(CellRange range)
        {
            if (range == null)
            {
                return "模板规则没有保存测量值区域。";
            }

            if (string.IsNullOrWhiteSpace(range.SheetName))
            {
                return "模板规则中的测量值区域缺少工作表名。";
            }

            if (range.StartRow <= 0 || range.StartColumn <= 0)
            {
                return $"模板规则中的测量值区域起点坐标非法：{range}。";
            }

            if (range.EndRow < range.StartRow || range.EndColumn < range.StartColumn)
            {
                return $"模板规则中的测量值区域终点小于起点：{range}。";
            }

            return $"模板规则中的测量值区域坐标非法：{range}。";
        }

        private static string BuildNoWritableCellReason(
            CellRange range,
            int totalCells,
            int missingCells,
            int mergedCoveredCells)
        {
            var details = new List<string>();
            if (totalCells <= 0)
            {
                details.Add("区域内没有单元格");
            }

            if (missingCells > 0)
            {
                details.Add($"{missingCells} 个单元格未被当前快照读取");
            }

            if (mergedCoveredCells > 0)
            {
                details.Add($"{mergedCoveredCells} 个单元格属于合并区域的非左上角，不能单独写入");
            }

            if (details.Count == 0)
            {
                details.Add("区域内没有可写入的普通单元格");
            }

            return $"测量值区域 {range} 未解析到可写入单元格（{string.Join("；", details)}）。";
        }

        private static string BuildCellKey(int row, int column)
        {
            return row + ":" + column;
        }
    }

    public sealed class WritableCellResolution
    {
        public List<CellAddress> Cells { get; set; } = new List<CellAddress>();
        public string FailureReason { get; set; }
    }
}
