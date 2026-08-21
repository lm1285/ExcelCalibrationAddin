using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelCalibrationAddin.Core.Models
{
    public sealed class AbsoluteAreaPosition
    {
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public List<List<string>> Formulas { get; set; } = new List<List<string>>();
    }

    public sealed class MultiAreaPosition
    {
        public int RowOffset { get; set; }
        public int ColumnOffset { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public List<List<string>> Formulas { get; set; } = new List<List<string>>();
    }

    public sealed class MultiAreaPositionTemplate
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public DateTime CreatedUtc { get; set; }
        public DateTime UpdatedUtc { get; set; }
        public string SourceWorkbookName { get; set; }
        public string SourceSheetName { get; set; }
        public int SourceAnchorRow { get; set; }
        public int SourceAnchorColumn { get; set; }
        public List<MultiAreaPosition> Areas { get; set; } = new List<MultiAreaPosition>();

        public static MultiAreaPositionTemplate Create(
            string name,
            IEnumerable<AbsoluteAreaPosition> absoluteAreas,
            string id = null,
            DateTime? createdUtc = null)
        {
            var normalizedName = (name ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedName))
            {
                throw new InvalidOperationException("位置模板名称不能为空。");
            }

            var areas = (absoluteAreas ?? Enumerable.Empty<AbsoluteAreaPosition>())
                .Where(item => item != null)
                .Select(CloneAbsolute)
                .ToList();
            ValidateAbsoluteAreas(areas);

            var anchorRow = areas.Min(item => item.StartRow);
            var anchorColumn = areas.Min(item => item.StartColumn);
            var now = DateTime.UtcNow;
            var template = new MultiAreaPositionTemplate
            {
                Id = string.IsNullOrWhiteSpace(id) ? Guid.NewGuid().ToString("N") : id.Trim(),
                Name = normalizedName,
                CreatedUtc = createdUtc ?? now,
                UpdatedUtc = now,
                Areas = areas
                    .Select(item => new MultiAreaPosition
                    {
                        RowOffset = item.StartRow - anchorRow,
                        ColumnOffset = item.StartColumn - anchorColumn,
                        RowCount = item.RowCount,
                        ColumnCount = item.ColumnCount
                    })
                    .OrderBy(item => item.RowOffset)
                    .ThenBy(item => item.ColumnOffset)
                    .ToList()
            };
            template.Validate();
            return template;
        }

        public IReadOnlyList<AbsoluteAreaPosition> Resolve(int anchorRow, int anchorColumn)
        {
            Validate();
            if (anchorRow <= 0 || anchorColumn <= 0)
            {
                throw new InvalidOperationException("区域锚点必须位于有效单元格内。");
            }

            return Areas.Select(item => new AbsoluteAreaPosition
            {
                StartRow = checked(anchorRow + item.RowOffset),
                StartColumn = checked(anchorColumn + item.ColumnOffset),
                RowCount = item.RowCount,
                ColumnCount = item.ColumnCount
            }).ToList();
        }

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(Id))
            {
                throw new InvalidOperationException("位置模板缺少标识。");
            }

            if (string.IsNullOrWhiteSpace(Name))
            {
                throw new InvalidOperationException("位置模板名称不能为空。");
            }

            if (Areas == null || Areas.Count == 0)
            {
                throw new InvalidOperationException("位置模板至少需要包含一个区域。");
            }

            foreach (var area in Areas)
            {
                if (area == null || area.RowOffset < 0 || area.ColumnOffset < 0 ||
                    area.RowCount <= 0 || area.ColumnCount <= 0)
                {
                    throw new InvalidOperationException("位置模板包含无效区域。");
                }

                if (area.Formulas != null && area.Formulas.Count > 0 &&
                    (area.Formulas.Count != area.RowCount || area.Formulas.Any(row => row == null || row.Count != area.ColumnCount)))
                {
                    throw new InvalidOperationException("位置模板中的公式矩阵与区域尺寸不一致。");
                }

            }

            EnsureNoOverlap(Areas.Select(item => new AbsoluteAreaPosition
            {
                StartRow = item.RowOffset + 1,
                StartColumn = item.ColumnOffset + 1,
                RowCount = item.RowCount,
                ColumnCount = item.ColumnCount
            }).ToList());
        }

        public MultiAreaPositionTemplate Clone()
        {
            return new MultiAreaPositionTemplate
            {
                Id = Id,
                Name = Name,
                CreatedUtc = CreatedUtc,
                UpdatedUtc = UpdatedUtc,
                SourceWorkbookName = SourceWorkbookName,
                SourceSheetName = SourceSheetName,
                SourceAnchorRow = SourceAnchorRow,
                SourceAnchorColumn = SourceAnchorColumn,
                Areas = (Areas ?? new List<MultiAreaPosition>()).Select(item => new MultiAreaPosition
                {
                    RowOffset = item.RowOffset,
                    ColumnOffset = item.ColumnOffset,
                    RowCount = item.RowCount,
                    ColumnCount = item.ColumnCount,
                    Formulas = (item.Formulas ?? new List<List<string>>()).Select(row => row?.ToList() ?? new List<string>()).ToList()
                }).ToList()
            };
        }

        private static AbsoluteAreaPosition CloneAbsolute(AbsoluteAreaPosition item)
        {
            return new AbsoluteAreaPosition
            {
                StartRow = item.StartRow,
                StartColumn = item.StartColumn,
                RowCount = item.RowCount,
                ColumnCount = item.ColumnCount,
                Formulas = (item.Formulas ?? new List<List<string>>()).Select(row => row?.ToList() ?? new List<string>()).ToList()
            };
        }

        private static void ValidateAbsoluteAreas(IReadOnlyList<AbsoluteAreaPosition> areas)
        {
            if (areas.Count == 0)
            {
                throw new InvalidOperationException("请先选择至少一个待复制区域。");
            }

            if (areas.Any(item => item.StartRow <= 0 || item.StartColumn <= 0 ||
                                  item.RowCount <= 0 || item.ColumnCount <= 0))
            {
                throw new InvalidOperationException("选区包含无效区域。");
            }

            EnsureNoOverlap(areas);
        }

        private static void EnsureNoOverlap(IReadOnlyList<AbsoluteAreaPosition> areas)
        {
            for (var i = 0; i < areas.Count; i++)
            {
                for (var j = i + 1; j < areas.Count; j++)
                {
                    if (Overlaps(areas[i], areas[j]))
                    {
                        throw new InvalidOperationException("位置模板中的区域不能互相重叠。");
                    }
                }
            }
        }

        private static bool Overlaps(AbsoluteAreaPosition left, AbsoluteAreaPosition right)
        {
            var leftEndRow = (long)left.StartRow + left.RowCount - 1;
            var leftEndColumn = (long)left.StartColumn + left.ColumnCount - 1;
            var rightEndRow = (long)right.StartRow + right.RowCount - 1;
            var rightEndColumn = (long)right.StartColumn + right.ColumnCount - 1;
            return left.StartRow <= rightEndRow && right.StartRow <= leftEndRow &&
                   left.StartColumn <= rightEndColumn && right.StartColumn <= leftEndColumn;
        }
    }
}
