using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleDraftBuilder
    {
        private static bool IsRepeatabilityProject(string projectName)
        {
            return NormalizeHeaderText(projectName).IndexOf("\u91CD\u590D\u6027", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsAverageAsErrorProject(string projectName)
        {
            return NormalizeHeaderText(projectName).IndexOf("\u54CD\u5E94\u65F6\u95F4", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsMeasurementSearch(string[] keywords)
        {
            return keywords == MeasurementKeywords;
        }

        private static CellRange FindErrorRangeByProjectTitle(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            string projectName,
            params CellRange[] excludedRanges)
        {
            var aliases = BuildErrorHeaderAliases(projectName);
            if (aliases.Count == 0)
            {
                return null;
            }

            var referenceDataStartRow = excludedRanges
                .Where(range => range != null)
                .Select(range => range.StartRow)
                .DefaultIfEmpty(startRow + 1)
                .Max();

            var searchEndRow = Math.Min(endRow, startRow + 6);
            ResultHeaderCandidate best = null;
            foreach (var cell in sheet.Cells
                .Where(item =>
                    item.Row >= startRow &&
                    item.Row <= searchEndRow &&
                    !string.IsNullOrWhiteSpace(item.Text) &&
                    !LooksLikeSectionTitle(item.Text) &&
                    IsErrorHeaderForProject(item.Text, aliases))
                .OrderBy(item => item.Row)
                .ThenBy(item => item.Column))
            {
                var columnStart = cell.MergeRange?.StartColumn ?? cell.Column;
                var columnEnd = cell.MergeRange?.EndColumn ?? cell.Column;
                if (excludedRanges.Any(range => IsColumnRangeOverlap(columnStart, columnEnd, range)))
                {
                    continue;
                }

                var headerBottomRow = cell.MergeRange?.EndRow ?? cell.Row;
                var candidateDataStartRow = Math.Max(headerBottomRow + 1, referenceDataStartRow);
                var dataStartRow = FindFirstFormulaRow(sheet, candidateDataStartRow, endRow, columnStart, columnEnd)
                    ?? FindFirstDataRow(sheet, candidateDataStartRow, endRow, columnStart, columnEnd);
                if (dataStartRow <= 0)
                {
                    continue;
                }

                var formulaCount = CountFormulaCells(sheet, dataStartRow, endRow, columnStart, columnEnd);
                var dataCount = CountDataCells(sheet, dataStartRow, endRow, columnStart, columnEnd);
                var score = ScoreErrorHeader(cell.Text, aliases) + formulaCount * 10 + dataCount * 2;
                if (best == null || score > best.Score)
                {
                    best = new ResultHeaderCandidate
                    {
                        StartColumn = columnStart,
                        EndColumn = columnEnd,
                        DataStartRow = dataStartRow,
                        Score = score
                    };
                }
            }

            if (best == null)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = best.DataStartRow,
                EndRow = endRow,
                StartColumn = best.StartColumn,
                EndColumn = best.EndColumn
            };
        }

        private static List<string> BuildErrorHeaderAliases(string projectName)
        {
            var normalized = NormalizeHeaderText(projectName);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return new List<string>();
            }

            var aliases = new List<string>
            {
                normalized,
                "\u8BEF\u5DEE",
                "\u793A\u503C\u8BEF\u5DEE",
                "\u76F8\u5BF9\u8BEF\u5DEE",
                "\u91CD\u590D\u6027"
            };
            if (normalized.IndexOf("\u91CD\u590D\u6027", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                aliases.Add("\u91CD\u590D\u6027");
            }
            if (normalized.IndexOf("\u7A33\u5B9A\u6027", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                aliases.Add("\u7A33\u5B9A\u6027");
            }
            if (normalized.IndexOf("\u8BEF\u5DEE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                aliases.Add("\u8BEF\u5DEE");
            }

            return aliases.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static List<MeasurementJudgementConstraint> BuildAdditionalJudgementConstraints(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            string projectName,
            HeaderBand headerBand,
            CellRange primaryErrorRange,
            CellRange primaryTechnicalRange,
            CellRange primaryResultRange,
            params CellRange[] occupiedRanges)
        {
            var occupied = new List<CellRange>(occupiedRanges ?? Array.Empty<CellRange>())
            {
                primaryErrorRange,
                primaryTechnicalRange,
                primaryResultRange
            };
            var extraErrorRanges = CollectErrorRangesByProjectTitle(
                    sheet,
                    startRow,
                    endRow,
                    projectName,
                    occupied.ToArray())
                .Where(range => !RangesOverlap(range, primaryErrorRange))
                .ToList();
            extraErrorRanges.AddRange(CollectDataRanges(
                    sheet,
                    startRow,
                    endRow,
                    ErrorKeywords,
                    occupied.Concat(extraErrorRanges).ToArray())
                .Where(range => extraErrorRanges.All(existing => !RangesOverlap(existing, range)) &&
                    !RangesOverlap(range, primaryErrorRange)));

            var extraTechnicalRanges = CollectDataRanges(
                    sheet,
                    startRow,
                    endRow,
                    TechnicalKeywords,
                    occupied.Concat(extraErrorRanges).ToArray())
                .Concat(CollectLayoutRanges(
                    sheet,
                    headerBand,
                    endRow,
                    TechnicalKeywords,
                    occupied.Concat(extraErrorRanges).ToArray()))
                .Where(range => !RangesOverlap(range, primaryTechnicalRange) &&
                    occupied.Concat(extraErrorRanges).All(existing => !RangesOverlap(existing, range)))
                .GroupBy(RangeKey)
                .Select(group => group.First())
                .ToList();
            var extraResultRanges = CollectDataRanges(
                    sheet,
                    startRow,
                    endRow,
                    ResultKeywords,
                    occupied.Concat(extraErrorRanges).Concat(extraTechnicalRanges).ToArray())
                .Concat(CollectLayoutRanges(
                    sheet,
                    headerBand,
                    endRow,
                    ResultKeywords,
                    occupied.Concat(extraErrorRanges).Concat(extraTechnicalRanges).ToArray()))
                .Where(range => !RangesOverlap(range, primaryResultRange))
                .GroupBy(RangeKey)
                .Select(group => group.First())
                .ToList();

            var count = Math.Max(extraErrorRanges.Count, Math.Max(extraTechnicalRanges.Count, extraResultRanges.Count));
            var constraints = new List<MeasurementJudgementConstraint>();
            for (var index = 0; index < count; index++)
            {
                var errorRange = index < extraErrorRanges.Count ? extraErrorRanges[index] : null;
                var technicalRange = index < extraTechnicalRanges.Count ? extraTechnicalRanges[index] : null;
                var resultRange = index < extraResultRanges.Count ? extraResultRanges[index] : null;
                if (errorRange == null && technicalRange == null && resultRange == null)
                {
                    continue;
                }

                constraints.Add(new MeasurementJudgementConstraint
                {
                    Name = "附属判定" + (index + 1),
                    ErrorSource = errorRange == null ? null : new ParameterSource { Name = "误差", Range = CloneRange(errorRange) },
                    MpeSource = technicalRange == null ? null : new ParameterSource { Name = "技术要求", Range = CloneRange(technicalRange) },
                    ResultSource = resultRange == null ? null : new ParameterSource { Name = "结论", Range = CloneRange(resultRange) }
                });
            }

            return constraints;
        }

        private static List<CellRange> CollectErrorRangesByProjectTitle(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            string projectName,
            params CellRange[] excludedRanges)
        {
            var aliases = BuildErrorHeaderAliases(projectName);
            if (aliases.Count == 0)
            {
                return new List<CellRange>();
            }

            var referenceDataStartRow = excludedRanges
                .Where(range => range != null)
                .Select(range => range.StartRow)
                .DefaultIfEmpty(startRow + 1)
                .Max();
            var searchEndRow = Math.Min(endRow, startRow + 6);
            var candidates = new List<ResultHeaderCandidate>();
            foreach (var cell in sheet.Cells
                .Where(item =>
                    item.Row >= startRow &&
                    item.Row <= searchEndRow &&
                    !string.IsNullOrWhiteSpace(item.Text) &&
                    !LooksLikeSectionTitle(item.Text) &&
                    IsErrorHeaderForProject(item.Text, aliases))
                .OrderBy(item => item.Row)
                .ThenBy(item => item.Column))
            {
                var columnStart = cell.MergeRange?.StartColumn ?? cell.Column;
                var columnEnd = cell.MergeRange?.EndColumn ?? cell.Column;
                if (excludedRanges.Any(range => IsColumnRangeOverlap(columnStart, columnEnd, range)) ||
                    candidates.Any(existing => columnStart <= existing.EndColumn && existing.StartColumn <= columnEnd))
                {
                    continue;
                }

                var headerBottomRow = cell.MergeRange?.EndRow ?? cell.Row;
                var candidateDataStartRow = Math.Max(headerBottomRow + 1, referenceDataStartRow);
                var dataStartRow = FindFirstFormulaRow(sheet, candidateDataStartRow, endRow, columnStart, columnEnd)
                    ?? FindFirstDataRow(sheet, candidateDataStartRow, endRow, columnStart, columnEnd);
                if (dataStartRow <= 0)
                {
                    continue;
                }

                var formulaCount = CountFormulaCells(sheet, dataStartRow, endRow, columnStart, columnEnd);
                var dataCount = CountDataCells(sheet, dataStartRow, endRow, columnStart, columnEnd);
                candidates.Add(new ResultHeaderCandidate
                {
                    StartColumn = columnStart,
                    EndColumn = columnEnd,
                    DataStartRow = dataStartRow,
                    Score = ScoreErrorHeader(cell.Text, aliases) + formulaCount * 10 + dataCount * 2
                });
            }

            return candidates
                .OrderByDescending(item => item.Score)
                .Select(item => new CellRange
                {
                    SheetName = sheet.Name,
                    StartRow = item.DataStartRow,
                    EndRow = endRow,
                    StartColumn = item.StartColumn,
                    EndColumn = item.EndColumn
                })
                .ToList();
        }

        private static List<CellRange> CollectDataRanges(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            string[] keywords,
            params CellRange[] excludedRanges)
        {
            var headerCandidates = sheet.Cells
                .Where(cell =>
                    cell.Row >= startRow &&
                    cell.Row <= Math.Min(endRow, startRow + 6) &&
                    !string.IsNullOrWhiteSpace(cell.Text) &&
                    !LooksLikeSectionTitle(cell.Text) &&
                    !LooksLikeWrongFieldHeader(cell.Text, keywords) &&
                    keywords.Any(keyword => MatchesKeyword(cell.Text, keyword)))
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();
            var ranges = new List<CellRange>();
            foreach (var candidate in headerCandidates)
            {
                var columnStart = candidate.MergeRange?.StartColumn ?? candidate.Column;
                var columnEnd = candidate.MergeRange?.EndColumn ?? candidate.Column;
                if (excludedRanges.Any(range => IsColumnRangeOverlap(columnStart, columnEnd, range)) ||
                    ranges.Any(range => IsColumnRangeOverlap(columnStart, columnEnd, range)))
                {
                    continue;
                }

                var range = BuildDataRangeFromHeader(sheet, startRow, endRow, keywords, candidate);
                if (range != null)
                {
                    ranges.Add(range);
                }
            }

            return ranges;
        }

        private static List<CellRange> CollectLayoutRanges(
            SheetSnapshot sheet,
            HeaderBand headerBand,
            int endRow,
            string[] keywords,
            params CellRange[] excludedRanges)
        {
            var ranges = new List<CellRange>();
            if (headerBand == null)
            {
                return ranges;
            }

            var remaining = excludedRanges?.Where(range => range != null).ToList() ?? new List<CellRange>();
            while (true)
            {
                var inferred = InferRangeFromLayout(sheet, headerBand, endRow, keywords, remaining.ToArray());
                if (inferred == null || remaining.Any(range => RangesOverlap(range, inferred)))
                {
                    break;
                }

                ranges.Add(inferred);
                remaining.Add(inferred);
            }

            return ranges;
        }

        private static CellRange BuildDataRangeFromHeader(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            string[] keywords,
            CellMeta candidate)
        {
            if (candidate == null)
            {
                return null;
            }

            var selectedHeader = SelectBestHeaderCandidate(sheet, startRow, endRow, keywords, new List<CellMeta> { candidate });
            if (selectedHeader == null)
            {
                return null;
            }

            var effectiveRange = RefineMeasurementColumns(
                sheet,
                selectedHeader.HeaderBottomRow,
                endRow,
                selectedHeader.StartColumn,
                selectedHeader.EndColumn,
                keywords);
            var dataStartRow = FindFirstDataRow(
                sheet,
                effectiveRange.HeaderBottomRow + 1,
                endRow,
                effectiveRange.StartColumn,
                effectiveRange.EndColumn);
            if (dataStartRow <= 0 && !HasSufficientDataBelow(
                sheet,
                effectiveRange.HeaderBottomRow,
                endRow,
                effectiveRange.StartColumn,
                effectiveRange.EndColumn))
            {
                return null;
            }

            if (dataStartRow <= 0)
            {
                dataStartRow = Math.Min(endRow, effectiveRange.HeaderBottomRow + 1);
            }

            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = dataStartRow,
                EndRow = endRow,
                StartColumn = effectiveRange.StartColumn,
                EndColumn = effectiveRange.EndColumn
            };
        }

        private static string RangeKey(CellRange range)
        {
            return range == null
                ? string.Empty
                : string.Join(":", new[]
                {
                    range.SheetName ?? string.Empty,
                    range.StartRow.ToString(),
                    range.StartColumn.ToString(),
                    range.EndRow.ToString(),
                    range.EndColumn.ToString()
                });
        }

        private static CellRange CloneRange(CellRange range)
        {
            return range == null
                ? null
                : new CellRange
                {
                    SheetName = range.SheetName,
                    StartRow = range.StartRow,
                    EndRow = range.EndRow,
                    StartColumn = range.StartColumn,
                    EndColumn = range.EndColumn
                };
        }

        private static bool IsErrorHeaderForProject(string text, IReadOnlyList<string> aliases)
        {
            var normalized = NormalizeHeaderText(text);
            if (string.IsNullOrWhiteSpace(normalized) ||
                normalized.IndexOf("\u6280\u672F\u8981\u6C42", StringComparison.OrdinalIgnoreCase) >= 0 ||
                normalized.IndexOf("\u7ED3\u8BBA", StringComparison.OrdinalIgnoreCase) >= 0 ||
                normalized.IndexOf("\u9650", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            return aliases.Any(alias =>
                string.Equals(normalized, alias, StringComparison.OrdinalIgnoreCase) ||
                normalized.IndexOf(alias, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static double ScoreErrorHeader(string text, IReadOnlyList<string> aliases)
        {
            var normalized = NormalizeHeaderText(text);
            var score = aliases.Any(alias => string.Equals(normalized, alias, StringComparison.OrdinalIgnoreCase)) ? 80 : 50;
            if (normalized.IndexOf("\u8BEF\u5DEE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                score += 10;
            }

            return score;
        }

        private static bool IsColumnRangeOverlap(int startColumn, int endColumn, CellRange range)
        {
            return range != null && startColumn <= range.EndColumn && range.StartColumn <= endColumn;
        }

        private static bool HasInlineParameterValue(string text, string[] keywords)
        {
            var value = text ?? string.Empty;
            if (!value.Contains(":") && !value.Contains("\uFF1A"))
            {
                return false;
            }

            var parts = value.Split(new[] { ':', '\uFF1A' }, 2);
            if (parts.Length < 2 || string.IsNullOrWhiteSpace(parts[1]))
            {
                return false;
            }

            return keywords.Any(keyword => MatchesKeyword(parts[0], keyword)) &&
                (SheetRowContentAnalyzer.LooksNumeric(parts[1]) || LooksLikeRangeExpression(parts[1]));
        }

        private static bool LooksLikeRangeExpression(string text)
        {
            var value = (text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return value.Any(char.IsDigit) &&
                   (value.Contains("~") ||
                    value.Contains("\uFF5E") ||
                    value.Contains("-") ||
                    value.IndexOf("\u81F3", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static CellMeta FindRightSideValueCell(SheetSnapshot sheet, CellMeta labelCell, int endRow)
        {
            var labelEndColumn = labelCell.MergeRange?.EndColumn ?? labelCell.Column;
            var searchEndColumn = labelEndColumn + 4;
            var candidates = sheet.Cells
                .Where(cell =>
                    cell.Row == labelCell.Row &&
                    cell.Column > labelEndColumn &&
                    cell.Column <= searchEndColumn &&
                    !string.IsNullOrWhiteSpace(cell.Text))
                .OrderBy(cell => cell.Column)
                .ToList();

            foreach (var candidate in candidates)
            {
            if (LooksLikeInlineNote(candidate.Text) || SheetRowContentAnalyzer.LooksNumeric(candidate.Text) || LooksLikeRangeExpression(candidate.Text))
                {
                    return candidate;
                }
            }

            return null;
        }

        private static CellRange BuildSingleCellRange(string sheetName, CellMeta cell)
        {
            var mergeRange = cell.MergeRange;
            return new CellRange
            {
                SheetName = sheetName,
                StartRow = mergeRange?.StartRow ?? cell.Row,
                EndRow = mergeRange?.EndRow ?? cell.Row,
                StartColumn = mergeRange?.StartColumn ?? cell.Column,
                EndColumn = mergeRange?.EndColumn ?? cell.Column
            };
        }

        private static HeaderCandidate SelectBestHeaderCandidate(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            string[] keywords,
            List<CellMeta> candidates)
        {
            HeaderCandidate best = null;

            foreach (var candidate in candidates)
            {
                var columnStart = candidate.MergeRange?.StartColumn ?? candidate.Column;
                var columnEnd = candidate.MergeRange?.EndColumn ?? candidate.Column;
                var headerBottomRow = candidate.MergeRange?.EndRow ?? candidate.Row;
                var text = (candidate.Text ?? string.Empty).Trim();
                var span = Math.Max(1, columnEnd - columnStart + 1);
                var peerHeaders = CountPeerHeaders(sheet, candidate.Row, startRow, endRow);
                var dataCells = CountDataCells(sheet, headerBottomRow + 1, endRow, columnStart, columnEnd);
                var firstDataRow = FindFirstDataRow(sheet, headerBottomRow + 1, endRow, columnStart, columnEnd);
                var score = 0d;

                score += Math.Max(0, 40 - (candidate.Row - startRow) * 4);
                score += Math.Min(24, peerHeaders * 6);
                score += Math.Min(18, dataCells * 3);
                score += keywords.Max(keyword => ScoreKeywordMatch(text, keyword));

                if (firstDataRow > 0)
                {
                    score += 18;
                }

                if (span <= 4)
                {
                    score += 8;
                }
                else
                {
                    score -= Math.Min(12, (span - 4) * 3);
                }

                if (candidate.Row == startRow && candidate.Column <= 2 && LooksLikeSectionTitle(text))
                {
                    score -= 80;
                }

                if (LooksLikeInlineNote(text))
                {
                    score -= 35;
                }

                if (dataCells == 0)
                {
                    score -= 45;
                }

                if (best == null || score > best.Score)
                {
                    best = new HeaderCandidate
                    {
                        StartColumn = columnStart,
                        EndColumn = columnEnd,
                        HeaderBottomRow = headerBottomRow,
                        Score = score
                    };
                }
            }

            return best != null && best.Score >= 20 ? best : null;
        }

        private static bool IsColumnInsideRange(int column, CellRange range)
        {
            return range != null && column >= range.StartColumn && column <= range.EndColumn;
        }

        private static bool RangesOverlap(CellRange left, CellRange right)
        {
            return left != null &&
                right != null &&
                string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase) &&
                left.StartColumn <= right.EndColumn &&
                right.StartColumn <= left.EndColumn;
        }

    }
}
