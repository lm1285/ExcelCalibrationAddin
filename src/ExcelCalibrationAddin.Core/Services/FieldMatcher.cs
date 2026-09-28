using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class FieldMatcher
    {
        // A multi-gas record can contain the same calibration item once per
        // gas. Keep a generous guard for malformed sheets, but do not cap
        // normal repeated item blocks at the first dozen.
        private const int MaxSectionsPerSheet = 100;

        private static readonly Regex NumericChildTitleRegex = new Regex(
            @"^\s*(?<number>(?>\d+(?:\s*\.\s*\d+)+))\s*(?:[、)）.,，\-:：]|\s+)\s*(?<title>.+?)\s*$",
            RegexOptions.Compiled);

        private static readonly Regex NumericTopLevelTitleRegex = new Regex(
            @"^\s*(?<number>\d+)\s*[、)）.,，\-:：]\s*(?<title>.+?)\s*$",
            RegexOptions.Compiled);

        private static readonly Regex ChineseTopLevelTitleRegex = new Regex(
            @"^\s*(?<number>[一二三四五六七八九十]+)\s*[、)）.,，\-:：]\s*(?<title>.+?)\s*$",
            RegexOptions.Compiled);

        private static readonly Regex NumericChildNumberOnlyRegex = new Regex(
            @"^\s*(?<number>(?>\d+(?:\s*\.\s*\d+)+))\s*[、)）.,，\-:：]?\s*$",
            RegexOptions.Compiled);

        private static readonly Regex TopLevelNumberOnlyRegex = new Regex(
            @"^\s*(?:\d+|[一二三四五六七八九十]+)\s*[、)）.,，\-:：]\s*$",
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
            var numberedFields = BuildNumberedSectionFields(sheet);
            if (numberedFields.Count > 0)
            {
                // Some books number an outer context heading (for example a gas
                // block) while the actual calibration items below it are unnumbered.
                // In that case the inner item headings are the useful fields.
                if (numberedFields.All(field => !field.Reason.Contains("按子标题识别")))
                {
                    var fallbackFields = BuildUnnumberedSectionFields(sheet);
                    if (fallbackFields.Count > 0) return fallbackFields;
                }
                return numberedFields;
            }

            var unnumberedFields = BuildUnnumberedSectionFields(sheet);
            return unnumberedFields.Count > 0 ? unnumberedFields : BuildHeaderFallbackFields(sheet);
        }

        public static bool IsNumberedSectionTitleText(string text)
        {
            return TryParseNumberedSectionTitle(text, out _, out _);
        }

        private static List<RecognizedField> BuildNumberedSectionFields(SheetSnapshot sheet)
        {
            var rows = sheet.Cells
                .GroupBy(cell => cell.Row)
                .OrderBy(group => group.Key)
                .ToList();

            var sectionMarkers = rows
                .Select(group => FindNumberedSectionMarker(group.ToList()))
                .Where(marker => marker != null)
                .GroupBy(marker => marker.Row)
                .Select(group => group.First())
                .OrderBy(marker => marker.Row)
                .ToList();

            if (sectionMarkers.Count == 0)
            {
                return new List<RecognizedField>();
            }

            var maxColumn = sheet.Cells.Count == 0 ? 1 : sheet.Cells.Max(cell => cell.Column);
            var fields = new List<RecognizedField>();
            var topLevelMarkers = sectionMarkers.Where(marker => marker.Level == 1).ToList();

            if (topLevelMarkers.Count == 0)
            {
                foreach (var marker in sectionMarkers)
                {
                    fields.Add(BuildSectionField(sheet, sectionMarkers, marker, maxColumn, "按子标题识别"));
                }

                return fields.Take(MaxSectionsPerSheet).ToList();
            }

            var firstTopLevelRow = topLevelMarkers[0].Row;
            foreach (var orphan in sectionMarkers.Where(marker => marker.Level > 1 && marker.Row < firstTopLevelRow))
            {
                fields.Add(BuildSectionField(sheet, sectionMarkers, orphan, maxColumn, "按子标题识别"));
            }

            for (var index = 0; index < topLevelMarkers.Count; index++)
            {
                var topLevel = topLevelMarkers[index];
                var nextTopLevelRow = index + 1 < topLevelMarkers.Count
                    ? topLevelMarkers[index + 1].Row
                    : int.MaxValue;
                var children = sectionMarkers
                    .Where(marker => marker.Level > 1 && marker.Row > topLevel.Row && marker.Row < nextTopLevelRow)
                    .ToList();

                if (children.Count == 0)
                {
                    fields.Add(BuildSectionField(sheet, sectionMarkers, topLevel, maxColumn, "按大标题识别"));
                    continue;
                }

                foreach (var child in children)
                {
                    fields.Add(BuildSectionField(sheet, sectionMarkers, child, maxColumn, "按子标题识别"));
                }
            }

            return fields
                .OrderBy(field => field.Range.StartRow)
                .Take(MaxSectionsPerSheet)
                .ToList();
        }

        private static List<RecognizedField> BuildUnnumberedSectionFields(SheetSnapshot sheet)
        {
            var markers = sheet.Cells
                .GroupBy(cell => cell.Row)
                .OrderBy(group => group.Key)
                .Select(group => FindUnnumberedSectionMarker(group.ToList()))
                .Where(marker => marker != null)
                .GroupBy(marker => marker.Row)
                .Select(group => group.First())
                .OrderBy(marker => marker.Row)
                .ToList();
            var maxColumn = sheet.Cells.Count == 0 ? 1 : sheet.Cells.Max(cell => cell.Column);

            return markers
                .Select(marker => BuildSectionField(sheet, markers, marker, maxColumn, "按无编号标题识别"))
                .Take(MaxSectionsPerSheet)
                .ToList();
        }

        private static RecognizedField BuildSectionField(
            SheetSnapshot sheet,
            IReadOnlyList<SectionMarker> allMarkers,
            SectionMarker marker,
            int maxColumn,
            string reasonPrefix)
        {
            var nextRow = allMarkers
                .Where(candidate => candidate.Row > marker.Row)
                .Select(candidate => candidate.Row)
                .DefaultIfEmpty(InferLastContentRowExcludingTrailingNotes(sheet, marker.Row) + 1)
                .Min();
            var alias = CleanSectionTitle(marker.Text);

            return new RecognizedField
            {
                Alias = alias,
                Score = ScoreSection(alias),
                Reason = $"{reasonPrefix}：{alias}",
                Range = new CellRange
                {
                    SheetName = sheet.Name,
                    StartRow = marker.Row,
                    EndRow = Math.Max(marker.Row, nextRow - 1),
                    StartColumn = 1,
                    EndColumn = maxColumn
                }
            };
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

        private static SectionMarker FindNumberedSectionMarker(List<CellMeta> rowCells)
        {
            foreach (var cell in rowCells.OrderBy(cell => cell.Column))
            {
                var text = (cell.Text ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(text) || cell.Column > 2)
                {
                    continue;
                }

                if (!TryParseNumberedSectionTitle(text, out var level, out _))
                {
                    continue;
                }

                return new SectionMarker
                {
                    Row = cell.Row,
                    Text = text,
                    Level = level
                };
            }

            return FindSplitNumberedSectionMarker(rowCells);
        }

        private static SectionMarker FindSplitNumberedSectionMarker(List<CellMeta> rowCells)
        {
            var textCells = (rowCells ?? new List<CellMeta>())
                .Where(cell => !string.IsNullOrWhiteSpace(cell?.Text))
                .GroupBy(cell => cell.MergeRange?.StartColumn ?? cell.Column)
                .Select(group => group.OrderBy(cell => cell.Column).First())
                .OrderBy(cell => cell.MergeRange?.StartColumn ?? cell.Column)
                .ToList();

            foreach (var numberCell in textCells)
            {
                var numberStartColumn = numberCell.MergeRange?.StartColumn ?? numberCell.Column;
                if (numberStartColumn > 2 || !LooksLikeStandaloneSectionNumber(numberCell.Text))
                {
                    continue;
                }

                var numberEndColumn = numberCell.MergeRange?.EndColumn ?? numberCell.Column;
                var titleCell = textCells.FirstOrDefault(cell =>
                {
                    var titleStartColumn = cell.MergeRange?.StartColumn ?? cell.Column;
                    return titleStartColumn > numberEndColumn &&
                        titleStartColumn <= numberEndColumn + 6 &&
                        LooksLikeTitleName((cell.Text ?? string.Empty).Trim());
                });
                if (titleCell == null)
                {
                    continue;
                }

                var combinedTitle = $"{numberCell.Text.Trim()} {titleCell.Text.Trim()}";
                if (!TryParseNumberedSectionTitle(combinedTitle, out var level, out _))
                {
                    continue;
                }

                return new SectionMarker
                {
                    Row = numberCell.Row,
                    Text = combinedTitle,
                    Level = level
                };
            }

            return null;
        }

        private static bool LooksLikeStandaloneSectionNumber(string text)
        {
            var value = (text ?? string.Empty).Trim().Replace('．', '.');
            return NumericChildNumberOnlyRegex.IsMatch(value) || TopLevelNumberOnlyRegex.IsMatch(value);
        }

        private static SectionMarker FindUnnumberedSectionMarker(List<CellMeta> rowCells)
        {
            foreach (var cell in rowCells.OrderBy(cell => cell.Column))
            {
                var text = (cell.Text ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(text) || cell.Column > 2 ||
                    !LooksLikeUnnumberedSectionTitle(text, rowCells))
                {
                    continue;
                }

                return new SectionMarker
                {
                    Row = cell.Row,
                    Text = text,
                    Level = 1
                };
            }

            return null;
        }

        private static bool TryParseNumberedSectionTitle(string text, out int level, out string title)
        {
            level = 0;
            title = string.Empty;
            var value = (text ?? string.Empty).Trim().Replace('．', '.');
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var childMatch = NumericChildTitleRegex.Match(value);
            if (childMatch.Success)
            {
                var number = Regex.Replace(childMatch.Groups["number"].Value, @"\s+", string.Empty);
                level = number.Count(ch => ch == '.') + 1;
                title = childMatch.Groups["title"].Value.Trim();
                return LooksLikeTitleName(title);
            }

            var topLevelMatch = NumericTopLevelTitleRegex.Match(value);
            if (!topLevelMatch.Success)
            {
                topLevelMatch = ChineseTopLevelTitleRegex.Match(value);
            }

            if (!topLevelMatch.Success)
            {
                return false;
            }

            level = 1;
            title = topLevelMatch.Groups["title"].Value.Trim();
            return LooksLikeTitleName(title);
        }

        private static bool LooksLikeTitleName(string title)
        {
            return !string.IsNullOrWhiteSpace(title) && title.Any(char.IsLetter);
        }

        private static bool LooksLikeUnnumberedSectionTitle(string text, List<CellMeta> rowCells)
        {
            var populatedCells = (rowCells ?? new List<CellMeta>())
                .Where(cell => !string.IsNullOrWhiteSpace(cell?.Text) || !string.IsNullOrWhiteSpace(cell?.Formula))
                .ToList();

            var isMergedTitle = (rowCells ?? new List<CellMeta>()).Any(cell =>
                string.Equals((cell?.Text ?? string.Empty).Trim(), (text ?? string.Empty).Trim(), StringComparison.Ordinal) &&
                cell.IsMerged && cell.MergeRange != null &&
                (cell.MergeRange.EndColumn > cell.MergeRange.StartColumn || cell.MergeRange.EndRow > cell.MergeRange.StartRow));

            return SectionKeywords.Any(keyword => text.Contains(keyword)) &&
                text.Length <= 24 &&
                (populatedCells.Count == 1 || isMergedTitle);
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
                : 80;
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
            public int Level { get; set; }
        }
    }
}
