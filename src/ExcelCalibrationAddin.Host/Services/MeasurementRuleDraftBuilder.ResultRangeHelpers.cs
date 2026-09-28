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
                    !(item.Row == startRow && string.Equals(NormalizeHeaderText(item.Text), NormalizeHeaderText(projectName), StringComparison.OrdinalIgnoreCase)) &&
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
                    var horizontal = BuildHorizontalDataRangeFromHeader(sheet, startRow, endRow, cell, columnStart, columnEnd);
                    if (horizontal != null)
                    {
                        var horizontalFormulaCount = CountFormulaCells(sheet, horizontal.StartRow, horizontal.EndRow, horizontal.StartColumn, horizontal.EndColumn);
                        var horizontalScore = ScoreErrorHeader(cell.Text, aliases) + horizontalFormulaCount * 10;
                        if (best == null || horizontalScore > best.Score)
                        {
                            best = new ResultHeaderCandidate
                            {
                                HeaderRow = cell.Row,
                                StartColumn = horizontal.StartColumn,
                                EndColumn = horizontal.EndColumn,
                                DataStartRow = horizontal.StartRow,
                                Score = horizontalScore
                            };
                        }
                    }
                    continue;
                }

                var formulaCount = CountFormulaCells(sheet, dataStartRow, endRow, columnStart, columnEnd);
                var dataCount = CountDataCells(sheet, dataStartRow, endRow, columnStart, columnEnd);
                var score = ScoreErrorHeader(cell.Text, aliases) + formulaCount * 10 + dataCount * 2;
                if (best == null || score > best.Score)
                {
                    best = new ResultHeaderCandidate
                    {
                        HeaderRow = cell.Row,
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

            var nextHeaderRow = sheet.Cells
                .Where(item => item.Row > best.HeaderRow && item.Row <= searchEndRow &&
                    item.Column >= best.StartColumn && item.Column <= best.EndColumn &&
                    IsErrorHeaderForProject(item.Text, aliases))
                .Select(item => item.Row)
                .DefaultIfEmpty(endRow + 1)
                .Min();
            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = best.DataStartRow,
                EndRow = Math.Min(endRow, nextHeaderRow - 1),
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
            out string pairingError,
            params CellRange[] occupiedRanges)
        {
            pairingError = string.Empty;
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
                primaryErrorRange)
                .Where(range => !SameRange(range, primaryErrorRange))
                .OrderBy(range => range.StartRow)
                .ThenBy(range => range.StartColumn)
                .ToList();
            extraErrorRanges.AddRange(CollectDataRanges(
                sheet,
                startRow,
                endRow,
                ErrorKeywords,
                new[] { primaryErrorRange }.Concat(extraErrorRanges).ToArray())
                .Where(range => extraErrorRanges.All(existing => !RangesOverlap(existing, range)) &&
                    !SameRange(range, primaryErrorRange)));
            extraErrorRanges = extraErrorRanges
                .OrderBy(range => range.StartRow)
                .ThenBy(range => range.StartColumn)
                .ToList();

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
                .Where(range => !SameRange(range, primaryTechnicalRange) &&
                    occupied.Concat(extraErrorRanges).All(existing => !SameRange(existing, range)))
                .GroupBy(RangeKey)
                .Select(group => group.First())
                .OrderBy(range => range.StartRow)
                .ThenBy(range => range.StartColumn)
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
                .Where(range => !SameRange(range, primaryResultRange))
                .GroupBy(RangeKey)
                .Select(group => group.First())
                .ToList();

            var pairedTechnicalRanges = PairRelatedRanges(extraErrorRanges, extraTechnicalRanges, out var technicalPairingError);
            var pairedResultRanges = PairRelatedRanges(extraErrorRanges, extraResultRanges, out var resultPairingError);
            var errors = new List<string>();
            var totalErrorRanges = (primaryErrorRange == null ? 0 : 1) + extraErrorRanges.Count;
            var totalTechnicalRanges = (primaryTechnicalRange == null ? 0 : 1) + extraTechnicalRanges.Count;
            var totalResultRanges = (primaryResultRange == null ? 0 : 1) + extraResultRanges.Count;
            if (totalErrorRanges != totalTechnicalRanges)
            {
                errors.Add($"同一校准项内误差区域({totalErrorRanges})与技术要求/MPE区域({totalTechnicalRanges})数量不一致");
            }
            if (!string.IsNullOrWhiteSpace(technicalPairingError)) errors.Add(technicalPairingError);
            if (!string.IsNullOrWhiteSpace(resultPairingError)) errors.Add(resultPairingError);
            if (totalResultRanges > 1 && totalResultRanges != totalErrorRanges)
            {
                errors.Add($"同一校准项内结论区域({totalResultRanges})与误差区域({totalErrorRanges})数量不一致");
            }
            pairingError = string.Join("；", errors.Distinct(StringComparer.Ordinal));

            var count = Math.Max(extraErrorRanges.Count, Math.Max(extraTechnicalRanges.Count, extraResultRanges.Count));
            var constraints = new List<MeasurementJudgementConstraint>();
            for (var index = 0; index < count; index++)
            {
                var errorRange = index < extraErrorRanges.Count ? extraErrorRanges[index] : null;
                var technicalRange = index < pairedTechnicalRanges.Count ? pairedTechnicalRanges[index] : null;
                var resultRange = index < pairedResultRanges.Count ? pairedResultRanges[index] : null;
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

        /// <summary>
        /// Pair related columns/rows by shared axis first and physical distance
        /// second. This works for the usual horizontal layout (same rows) and
        /// for vertically stacked judgement rows (same columns), while exposing
        /// ties instead of silently choosing one MPE.
        /// </summary>
        private static List<CellRange> PairRelatedRanges(
            IReadOnlyList<CellRange> anchors,
            IReadOnlyList<CellRange> candidates,
            out string pairingError)
        {
            pairingError = string.Empty;
            var result = new List<CellRange>();
            var used = new HashSet<int>();
            for (var anchorIndex = 0; anchorIndex < (anchors?.Count ?? 0); anchorIndex++)
            {
                var anchor = anchors[anchorIndex];
                var ranked = (candidates ?? Array.Empty<CellRange>())
                    .Select((candidate, index) => new { candidate, index, score = RangePairScore(anchor, candidate) })
                    .Where(item => item.candidate != null && !used.Contains(item.index))
                    .OrderByDescending(item => item.score)
                    .ThenBy(item => item.index)
                    .ToList();
                if (ranked.Count == 0)
                {
                    continue;
                }

                var best = ranked[0];
                var tied = ranked.Skip(1).Any(item => Math.Abs(item.score - best.score) < 0.0001);
                if (tied)
                {
                    pairingError = "误差与技术要求/MPE区域存在无法消除的配对歧义";
                    continue;
                }

                used.Add(best.index);
                result.Add(best.candidate);
            }

            // Preserve candidate count/order for an explicit count mismatch
            // diagnostic. Unpaired candidates remain visible as null slots in
            // the constraint list created by the caller.
            for (var index = 0; index < (candidates?.Count ?? 0); index++)
            {
                if (!used.Contains(index)) result.Add(candidates[index]);
            }
            return result;
        }

        private static double RangePairScore(CellRange left, CellRange right)
        {
            if (left == null || right == null) return double.MinValue;
            var rowOverlap = Math.Max(0, Math.Min(left.EndRow, right.EndRow) - Math.Max(left.StartRow, right.StartRow) + 1);
            var columnOverlap = Math.Max(0, Math.Min(left.EndColumn, right.EndColumn) - Math.Max(left.StartColumn, right.StartColumn) + 1);
            var rowDistance = Math.Abs(((left.StartRow + left.EndRow) / 2.0) - ((right.StartRow + right.EndRow) / 2.0));
            var columnDistance = Math.Abs(((left.StartColumn + left.EndColumn) / 2.0) - ((right.StartColumn + right.EndColumn) / 2.0));
            return rowOverlap * 100000d + columnOverlap * 100000d - rowDistance * 10d - columnDistance;
        }

        private static bool SameRange(CellRange left, CellRange right)
        {
            return left != null && right != null &&
                string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase) &&
                left.StartRow == right.StartRow && left.EndRow == right.EndRow &&
                left.StartColumn == right.StartColumn && left.EndColumn == right.EndColumn;
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
                    !(item.Row == startRow && string.Equals(NormalizeHeaderText(item.Text), NormalizeHeaderText(projectName), StringComparison.OrdinalIgnoreCase)) &&
                    IsErrorHeaderForProject(item.Text, aliases))
                .OrderBy(item => item.Row)
                .ThenBy(item => item.Column))
            {
                var columnStart = cell.MergeRange?.StartColumn ?? cell.Column;
                var columnEnd = cell.MergeRange?.EndColumn ?? cell.Column;
                if (excludedRanges.Any(range => IsColumnRangeOverlap(columnStart, columnEnd, range) &&
                    cell.Row < range.StartRow))
                {
                    continue;
                }

                var headerBottomRow = cell.MergeRange?.EndRow ?? cell.Row;
                var candidateDataStartRow = Math.Max(headerBottomRow + 1, referenceDataStartRow);
                var dataStartRow = FindFirstFormulaRow(sheet, candidateDataStartRow, endRow, columnStart, columnEnd)
                    ?? FindFirstDataRow(sheet, candidateDataStartRow, endRow, columnStart, columnEnd);
                if (dataStartRow <= 0)
                {
                    var horizontal = BuildHorizontalDataRangeFromHeader(sheet, startRow, endRow, cell, columnStart, columnEnd);
                    if (horizontal != null)
                    {
                        candidates.Add(new ResultHeaderCandidate
                        {
                            HeaderRow = cell.Row,
                            StartColumn = horizontal.StartColumn,
                            EndColumn = horizontal.EndColumn,
                            DataStartRow = horizontal.StartRow,
                            Score = ScoreErrorHeader(cell.Text, aliases) + CountFormulaCells(sheet, horizontal.StartRow, horizontal.EndRow, horizontal.StartColumn, horizontal.EndColumn) * 10
                        });
                    }
                    continue;
                }

                var formulaCount = CountFormulaCells(sheet, dataStartRow, endRow, columnStart, columnEnd);
                var dataCount = CountDataCells(sheet, dataStartRow, endRow, columnStart, columnEnd);
                candidates.Add(new ResultHeaderCandidate
                {
                    HeaderRow = cell.Row,
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
                    EndRow = candidates
                        .Where(next => next.HeaderRow > item.HeaderRow &&
                                       next.StartColumn <= item.EndColumn &&
                                       item.StartColumn <= next.EndColumn)
                        .Select(next => next.HeaderRow - 1)
                        .DefaultIfEmpty(endRow)
                        .Min(),
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
                    !(keywords == ErrorKeywords && cell.Row == startRow) &&
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
                if (excludedRanges.Any(range => IsColumnRangeOverlap(columnStart, columnEnd, range) &&
                                                candidate.Row < range.StartRow) ||
                    ranges.Any(range => IsColumnRangeOverlap(columnStart, columnEnd, range) &&
                                        candidate.Row < range.StartRow))
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

        private static CellRange BuildHorizontalDataRangeFromHeader(
            SheetSnapshot sheet,
            int startRow,
            int endRow,
            CellMeta header,
            int headerStartColumn,
            int headerEndColumn)
        {
            if (sheet == null || header == null)
            {
                return null;
            }

            var maxColumn = InferMaxColumn(sheet);
            var dataStartColumn = 0;
            var dataEndColumn = 0;
            var rowStart = Math.Max(startRow, header.MergeRange?.StartRow ?? header.Row);
            var rowEnd = Math.Min(endRow, Math.Max(rowStart, header.MergeRange?.EndRow ?? header.Row));
            for (var column = headerEndColumn + 1; column <= maxColumn; column++)
            {
                var hasData = sheet.Cells.Any(cell =>
                    cell.Column == column &&
                    cell.Row >= rowStart &&
                    cell.Row <= rowEnd &&
                    (SheetRowContentAnalyzer.LooksNumeric(cell.Text) || !string.IsNullOrWhiteSpace(cell.Formula)));
                if (!hasData)
                {
                    if (dataStartColumn > 0) break;
                    continue;
                }

                if (dataStartColumn == 0) dataStartColumn = column;
                dataEndColumn = column;
            }

            if (dataStartColumn <= 0 || dataEndColumn < dataStartColumn)
            {
                return null;
            }

            // Horizontal blocks usually keep one logical row per field. If the
            // first populated row is below a merged header, include that row.
            var populatedRows = sheet.Cells
                .Where(cell => cell.Column >= dataStartColumn && cell.Column <= dataEndColumn &&
                              cell.Row >= rowStart && cell.Row <= rowEnd &&
                              (SheetRowContentAnalyzer.LooksNumeric(cell.Text) || !string.IsNullOrWhiteSpace(cell.Formula)))
                .Select(cell => cell.Row)
                .Distinct()
                .OrderBy(value => value)
                .ToList();
            if (populatedRows.Count == 0)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = sheet.Name,
                StartRow = populatedRows.First(),
                EndRow = populatedRows.Last(),
                StartColumn = dataStartColumn,
                EndColumn = dataEndColumn
            };
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
            if (dataStartRow <= 0)
            {
                var horizontal = BuildHorizontalDataRangeFromHeader(sheet, startRow, endRow, candidate,
                    selectedHeader.StartColumn, selectedHeader.EndColumn);
                if (horizontal != null)
                {
                    return horizontal;
                }
            }

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
