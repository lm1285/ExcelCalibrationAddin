using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public sealed class SampleDataCaptureResult
    {
        public long VersionId { get; set; }
        public int SavedItemCount { get; set; }
        public List<string> SkippedItems { get; set; } = new List<string>();
    }

    public sealed class SampleDataUseCase
    {
        private static readonly Regex DecimalRegex = new Regex(@"[.,](?<digits>\d+)", RegexOptions.Compiled);
        private readonly IWorkbookSnapshotProvider _snapshotProvider;
        private readonly LocalTemplateRuleCacheRepository _repository;

        public SampleDataUseCase(IWorkbookSnapshotProvider snapshotProvider, LocalTemplateRuleCacheRepository repository)
        {
            _snapshotProvider = snapshotProvider ?? throw new ArgumentNullException(nameof(snapshotProvider));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public SampleDataCaptureResult CaptureAndSave(string templateFingerprint, IReadOnlyList<MeasurementRule> rules, ISet<string> selectedNames = null)
        {
            if (string.IsNullOrWhiteSpace(templateFingerprint)) throw new InvalidOperationException("当前没有有效的模板指纹。");
            var selectedRules = (rules ?? new List<MeasurementRule>()).Where(rule => rule != null &&
                (selectedNames == null || selectedNames.Count == 0 || selectedNames.Contains(ResolveName(rule)))).ToList();
            var ranges = selectedRules.SelectMany(rule => new[] { rule.TargetRange, rule.StandardValueSource?.Range }).Where(IsValidRange).ToList();
            var snapshot = _snapshotProvider.Capture(ranges);
            var items = new List<TemplateSampleData>();
            var skipped = new List<string>();
            foreach (var rule in selectedRules)
            {
                var item = CaptureRule(snapshot, rule, out var reason);
                if (item == null) skipped.Add($"{ResolveName(rule)}：{reason}"); else items.Add(item);
            }
            if (items.Count == 0) return new SampleDataCaptureResult { SkippedItems = skipped };
            var versionId = _repository.SaveSampleDataVersion(templateFingerprint, items);
            return new SampleDataCaptureResult { VersionId = versionId, SavedItemCount = items.Count, SkippedItems = skipped };
        }

        private static TemplateSampleData CaptureRule(WorkbookSnapshot snapshot, MeasurementRule rule, out string reason)
        {
            reason = "测量区为空或没有合法数值";
            var target = rule.TargetRange;
            var sheet = snapshot?.Sheets.FirstOrDefault(item => string.Equals(item.Name, target?.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null || target == null) return null;
            var cells = sheet.Cells.Where(cell => cell != null && cell.Row >= target.StartRow && cell.Row <= target.EndRow && cell.Column >= target.StartColumn && cell.Column <= target.EndColumn).ToList();
            var standardCells = ResolveStandardCells(snapshot, rule.StandardValueSource?.Range);
            var item = new TemplateSampleData { CalibrationItemName = ResolveName(rule), CalibrationItemKey = ResolveName(rule) };
            foreach (var row in cells.GroupBy(cell => cell.Row).OrderBy(group => group.Key))
            {
                var values = row.Select(ParseDouble).Where(value => value.HasValue).Select(value => value.Value).Take(30).ToList();
                if (values.Count == 0) continue;
                var first = row.First();
                item.Points.Add(new SampleDataPoint
                {
                    PointIndex = item.Points.Count + 1,
                    SourceRow = first.Row,
                    SourceColumn = first.Column,
                    StandardValue = standardCells.TryGetValue(first.Row, out var standard) ? standard : rule.FixedStandardValue,
                    MeasurementValues = values,
                    DecimalPlaces = row.Select(ResolveDecimalPlaces).DefaultIfEmpty(0).Max()
                });
            }
            return item.Points.Count == 0 ? null : item;
        }

        private static Dictionary<int, double> ResolveStandardCells(WorkbookSnapshot snapshot, CellRange range)
        {
            var result = new Dictionary<int, double>();
            var sheet = snapshot?.Sheets.FirstOrDefault(item => string.Equals(item.Name, range?.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null || range == null) return result;
            foreach (var cell in sheet.Cells.Where(cell => cell.Row >= range.StartRow && cell.Row <= range.EndRow && cell.Column >= range.StartColumn && cell.Column <= range.EndColumn).OrderBy(cell => cell.Row))
            {
                var value = ParseDouble(cell);
                if (value.HasValue && !result.ContainsKey(cell.Row)) result[cell.Row] = value.Value;
            }
            return result;
        }

        private static double? ParseDouble(CellMeta cell)
        {
            if (cell == null || !string.IsNullOrWhiteSpace(cell.Formula)) return null;
            double value;
            return double.TryParse(cell.RawValueText, NumberStyles.Float, CultureInfo.InvariantCulture, out value) && !double.IsNaN(value) && !double.IsInfinity(value) ? value : (double?)null;
        }

        private static int ResolveDecimalPlaces(CellMeta cell)
        {
            var format = cell?.NumberFormat ?? string.Empty;
            var match = DecimalRegex.Match(format);
            if (match.Success) return Math.Min(15, match.Groups["digits"].Value.Count(ch => ch == '0' || ch == '#'));
            var text = cell?.DisplayText;
            match = DecimalRegex.Match(text ?? string.Empty);
            return match.Success ? Math.Min(15, match.Groups["digits"].Value.Length) : 0;
        }

        private static string ResolveName(MeasurementRule rule) => string.IsNullOrWhiteSpace(rule?.FieldAlias) ? rule?.FieldName ?? string.Empty : rule.FieldAlias;
        private static bool IsValidRange(CellRange range) => range != null && range.StartRow > 0 && range.StartColumn > 0 && range.EndRow >= range.StartRow && range.EndColumn >= range.StartColumn && !string.IsNullOrWhiteSpace(range.SheetName);
    }
}
