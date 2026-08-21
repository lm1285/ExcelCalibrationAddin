using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class FieldMatcher
    {
        private const int MaxSectionsPerSheet = 12;

        private static readonly Regex SectionTitleRegex = new Regex(
            @"^\s*[一二三四五六七八九十\d]+\s*[、.)．\-]\s*.+",
            RegexOptions.Compiled);

        private static readonly string[] MeasurementKeywords =
        {
            "外观",
            "报警",
            "示值误差",
            "重复性",
            "响应时间",
            "漂移",
            "平均值",
            "均值",
            "误差",
            "不确定度",
            "量程"
        };

        private static readonly string[] SectionKeywords =
        {
            "外观", "报警功能", "报警动作值", "示值误差", "基本误差", "引用误差",
            "重复性", "稳定性", "漂移", "响应时间", "绝缘电阻", "零点漂移", "量程漂移"
        };

        public List<RecognizedField> MatchMeasurementFields(SheetSnapshot sheet)
        {
            var sectionFields = BuildSectionFields(sheet);
            return sectionFields.Count > 0 ? sectionFields : BuildHeaderFallbackFields(sheet);
        }

        private static List<RecognizedField> BuildSectionFields(SheetSnapshot sheet)
        {
            var rows = sheet.Cells
                .GroupBy(cell => cell.Row)
                .OrderBy(group => group.Key)
                .ToList();

            var sectionMarkers = rows
                .Select(group => FindSectionMarker(group.ToList()))
                .Where(marker => marker != null)
                .GroupBy(marker => marker.Row)
                .Select(group => group.First())
                .OrderBy(marker => marker.Row)
                .Take(MaxSectionsPerSheet)
                .ToList();

            if (sectionMarkers.Count == 0)
            {
                return new List<RecognizedField>();
            }

            var maxColumn = sheet.Cells.Count == 0 ? 1 : sheet.Cells.Max(cell => cell.Column);
            var fields = new List<RecognizedField>();

            for (var index = 0; index < sectionMarkers.Count; index++)
            {
                var marker = sectionMarkers[index];
                var nextRow = index + 1 < sectionMarkers.Count
                    ? sectionMarkers[index + 1].Row
                    : InferLastContentRowExcludingTrailingNotes(sheet, marker.Row) + 1;
                var sectionEndRow = Math.Max(marker.Row, nextRow - 1);
                var alias = CleanSectionTitle(marker.Text);

                fields.Add(new RecognizedField
                {
                    Alias = alias,
                    Score = ScoreSection(alias),
                    Reason = $"按项目块识别：{alias}",
                    Range = new CellRange
                    {
                        SheetName = sheet.Name,
                        StartRow = marker.Row,
                        EndRow = sectionEndRow,
                        StartColumn = 1,
                        EndColumn = maxColumn
                    }
                });
            }

            return fields;
        }

        private static List<RecognizedField> BuildHeaderFallbackFields(SheetSnapshot sheet)
        {
            return sheet.Headers
                .Select(header => string.Join("/",
                    header.Levels
                        .Select(item => (item ?? string.Empty).Trim())
                        .Where(item => !string.IsNullOrWhiteSpace(item))
                        .Distinct()))
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(12)
                .Select((text, index) => new RecognizedField
                {
                    Alias = text,
                    Score = ScoreSection(text),
                    Reason = $"按表头识别：{text}",
                    Range = new CellRange
                    {
                        SheetName = sheet.Name,
                        StartRow = 1,
                        EndRow = 6,
                        StartColumn = index + 1,
                        EndColumn = index + 1
                    }
                })
                .Where(item => item.Score >= 60)
                .ToList();
        }

        private static SectionMarker FindSectionMarker(List<CellMeta> rowCells)
        {
            foreach (var cell in rowCells.OrderBy(cell => cell.Column))
            {
                var text = (cell.Text ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(text) || cell.Column > 2)
                {
                    continue;
                }

                if (!SectionTitleRegex.IsMatch(text) && !LooksLikeUnnumberedSectionTitle(text, rowCells))
                {
                    continue;
                }

                if (!SectionKeywords.Any(keyword => text.Contains(keyword)))
                {
                    continue;
                }

                return new SectionMarker
                {
                    Row = cell.Row,
                    Text = text
                };
            }

            return null;
        }

        private static bool LooksLikeUnnumberedSectionTitle(string text, List<CellMeta> rowCells)
        {
            var populatedCells = (rowCells ?? new List<CellMeta>())
                .Where(cell => !string.IsNullOrWhiteSpace(cell?.Text) || !string.IsNullOrWhiteSpace(cell?.Formula))
                .ToList();

            return SectionKeywords.Any(keyword => text.Contains(keyword)) &&
                text.Length <= 24 &&
                populatedCells.Count == 1;
        }

        private static string CleanSectionTitle(string text)
        {
            var value = (text ?? string.Empty).Trim();
            value = Regex.Replace(value, @"[:：\s]*$", string.Empty);
            value = Regex.Replace(value, @"\s+", string.Empty);
            return value;
        }

        private static double ScoreSection(string alias)
        {
            return SectionKeywords.Any(keyword => alias.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                ? 96
                : 72;
        }

        private static int InferLastContentRowExcludingTrailingNotes(SheetSnapshot sheet, int sectionStartRow)
        {
            var contentRows = new SortedSet<int>();
            foreach (var cell in sheet.Cells)
            {
                if (!string.IsNullOrWhiteSpace(cell.Text) || !string.IsNullOrWhiteSpace(cell.Formula))
                {
                    contentRows.Add(cell.Row);
                }
            }

            while (contentRows.Count > 0)
            {
                var maxRow = contentRows.Max;
                if (maxRow < sectionStartRow || !IsTrailingNoteRow(sheet, maxRow))
                {
                    return Math.Max(sectionStartRow, maxRow);
                }

                contentRows.Remove(maxRow);
            }

            return sectionStartRow;
        }

        private static bool IsTrailingNoteRow(SheetSnapshot sheet, int row)
        {
            var rowTexts = sheet.Cells
                .Where(cell => cell.Row == row && !string.IsNullOrWhiteSpace(cell.Text))
                .OrderBy(cell => cell.Column)
                .Select(cell => new
                {
                    cell.Column,
                    Text = (cell.Text ?? string.Empty).Trim()
                })
                .ToList();

            if (rowTexts.Count == 0 || rowTexts[0].Column > 2)
            {
                return false;
            }

            var firstText = rowTexts[0].Text;
            return firstText.StartsWith("\u5907\u6CE8", StringComparison.OrdinalIgnoreCase) ||
                firstText.StartsWith("\u6CE8\u91CA", StringComparison.OrdinalIgnoreCase) ||
                firstText.StartsWith("\u8BF4\u660E", StringComparison.OrdinalIgnoreCase) ||
                firstText.StartsWith("\u6CE8\uFF1A", StringComparison.OrdinalIgnoreCase) ||
                firstText.StartsWith("\u6CE8:", StringComparison.OrdinalIgnoreCase);
        }

        private sealed class SectionMarker
        {
            public int Row { get; set; }
            public string Text { get; set; } = string.Empty;
        }
    }
}
