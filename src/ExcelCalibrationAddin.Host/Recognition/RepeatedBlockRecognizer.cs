using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Recognition
{
    /// <summary>
    /// Finds repeated, numbered vertical sections. A block title is context only;
    /// its text is deliberately excluded from the block signature.
    /// </summary>
    public static class RepeatedBlockRecognizer
    {
        private static readonly Regex TopLevelMarker = new Regex(
            @"^\s*(?<number>\d+)(?!\s*\.\s*\d)\s*(?<punct>[、)）:：])?\s*(?<title>.*)$",
            RegexOptions.Compiled);

        public static void Apply(SheetSnapshot sheet, IList<TemplateRegionMapping> mappings)
        {
            if (sheet == null || mappings == null || mappings.Count == 0) return;
            var markers = FindMarkers(sheet);
            var starts = FindBestRegularRun(markers);
            if (starts.Count < 2) return;

            var maxRow = sheet.Cells
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Text) || !string.IsNullOrWhiteSpace(cell.Formula))
                .Select(cell => cell.Row)
                .DefaultIfEmpty(starts.Last())
                .Max();
            var blockMappings = starts.Select((start, index) =>
            {
                var end = index + 1 < starts.Count ? starts[index + 1] - 1 : maxRow;
                var items = mappings.Where(mapping => mapping?.SectionRange != null &&
                    mapping.SectionRange.StartRow >= start && mapping.SectionRange.StartRow <= end)
                    .OrderBy(mapping => mapping.SectionRange.StartRow)
                    .ThenBy(mapping => mapping.SectionRange.StartColumn)
                    .ToList();
                return new { Start = start, End = end, Items = items };
            }).Where(block => block.Items.Count > 0).ToList();

            if (blockMappings.Count < 2) return;
            var signatures = blockMappings.Select(block => BuildSignature(sheet, block.Start, block.Items)).ToList();
            for (var blockIndex = 0; blockIndex < blockMappings.Count; blockIndex++)
            {
                var block = blockMappings[blockIndex];
                var range = new CellRange
                {
                    SheetName = sheet.Name,
                    StartRow = block.Start,
                    EndRow = block.End,
                    StartColumn = 1,
                    EndColumn = sheet.Cells.Select(cell => cell.Column).DefaultIfEmpty(1).Max()
                };
                for (var itemIndex = 0; itemIndex < block.Items.Count; itemIndex++)
                {
                    var mapping = block.Items[itemIndex];
                    mapping.BlockOrdinal = blockIndex + 1;
                    mapping.BlockRuleOrdinal = itemIndex + 1;
                    mapping.BlockRange = CloneRange(range);
                    mapping.BlockStructureSignature = signatures[blockIndex];
                    mapping.BlockItemStructureSignature = BuildSignature(sheet, block.Start, new[] { mapping });
                }
            }
        }

        private static List<int> FindMarkers(SheetSnapshot sheet)
        {
            return sheet.Cells.Where(cell => cell.Column <= 2 && !string.IsNullOrWhiteSpace(cell.Text))
                .GroupBy(cell => cell.Row)
                .Select(group => group.OrderBy(cell => cell.Column).First())
                .Where(cell =>
                {
                    var match = TopLevelMarker.Match((cell.Text ?? string.Empty).Trim());
                    if (!match.Success) return false;
                    var title = match.Groups["title"].Value.Trim();
                    // A bare ordinal is accepted only when explicitly punctuated (e.g. "2、").
                    return title.Length > 0 || match.Groups["punct"].Success;
                })
                .Select(cell => cell.Row)
                .Distinct()
                .OrderBy(row => row)
                .ToList();
        }

        private static List<int> FindBestRegularRun(IReadOnlyList<int> markers)
        {
            List<int> best = new List<int>();
            for (var first = 0; first < markers.Count; first++)
            {
                for (var second = first + 1; second < markers.Count; second++)
                {
                    var step = markers[second] - markers[first];
                    var run = new List<int> { markers[first], markers[second] };
                    var expected = markers[second] + step;
                    for (var next = second + 1; next < markers.Count; next++)
                    {
                        if (markers[next] == expected)
                        {
                            run.Add(markers[next]);
                            expected += step;
                        }
                        else if (markers[next] > expected)
                        {
                            break;
                        }
                    }
                    if (run.Count > best.Count || (run.Count == best.Count && run.Count > 0 && step < best[1] - best[0]))
                        best = run;
                }
            }
            return best;
        }

        private static string BuildSignature(SheetSnapshot sheet, int blockStart, IReadOnlyList<TemplateRegionMapping> items)
        {
            var regions = items.SelectMany(item => new[]
            {
                item.StandardValueRange, item.MeasurementValueRange, item.AverageValueRange,
                item.ErrorValueRange, item.TechnicalRequirementRange, item.UncertaintyRange,
                item.ResultRange
            }).Where(range => range != null).ToList();
            var cellShape = sheet.Cells.Where(cell => regions.Any(range =>
                    cell.Row >= range.StartRow && cell.Row <= range.EndRow &&
                    cell.Column >= range.StartColumn && cell.Column <= range.EndColumn))
                .Select(cell => string.Join(",", cell.Row - blockStart, cell.Column,
                    string.IsNullOrWhiteSpace(cell.FormulaR1C1) ? "V" : "F",
                    Normalize(cell.FormulaR1C1), Normalize(cell.NumberFormat)))
                .OrderBy(value => value, StringComparer.Ordinal);
            var shape = string.Join("|", items.Select(item => string.Join(":", new[]
            {
                NormalizeProjectName(item.ProjectName), Offset(item.SectionRange, blockStart),
                Offset(item.StandardValueRange, blockStart) + "@" + UnitSignature(sheet, item.StandardValueRange),
                Offset(item.MeasurementValueRange, blockStart) + "@" + UnitSignature(sheet, item.MeasurementValueRange),
                Offset(item.AverageValueRange, blockStart) + "@" + UnitSignature(sheet, item.AverageValueRange),
                Offset(item.ErrorValueRange, blockStart) + "@" + UnitSignature(sheet, item.ErrorValueRange),
                Offset(item.TechnicalRequirementRange, blockStart) + "@" + UnitSignature(sheet, item.TechnicalRequirementRange),
                Offset(item.UncertaintyRange, blockStart) + "@" + UnitSignature(sheet, item.UncertaintyRange),
                Offset(item.ResultRange, blockStart) + "@" + UnitSignature(sheet, item.ResultRange)
            })) + "#" + string.Join("|", cellShape));
            using (var sha = SHA256.Create())
                return string.Concat(sha.ComputeHash(Encoding.UTF8.GetBytes(shape)).Select(value => value.ToString("x2")));
        }

        private static string Offset(CellRange range, int blockStart)
        {
            return range == null ? "-" : string.Join(",", range.StartRow - blockStart,
                range.EndRow - blockStart, range.StartColumn, range.EndColumn);
        }

        private static string UnitSignature(SheetSnapshot sheet, CellRange range)
        {
            if (range == null) return string.Empty;
            var units = sheet.Cells.Where(cell => cell.Row >= range.StartRow && cell.Row <= range.EndRow &&
                    cell.Column >= range.StartColumn && cell.Column <= range.EndColumn)
                .Select(cell => ExcelCalibrationAddin.Host.Services.TemplateUnitParser.Extract(
                    cell.DisplayText, cell.Text, cell.NumberFormat))
                .Where(unit => !string.IsNullOrWhiteSpace(unit))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(unit => unit, StringComparer.OrdinalIgnoreCase);
            return string.Join(",", units);
        }

        private static string Normalize(string value)
        {
            return Regex.Replace((value ?? string.Empty).Trim().ToUpperInvariant(), @"\s+", string.Empty);
        }

        private static string NormalizeProjectName(string value)
        {
            var normalized = Normalize(value);
            return Regex.Replace(normalized, @"^\d+(?:\.\d+)*[、)）.,，:：\-]*", string.Empty);
        }

        private static CellRange CloneRange(CellRange range)
        {
            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                EndRow = range.EndRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }
    }
}
