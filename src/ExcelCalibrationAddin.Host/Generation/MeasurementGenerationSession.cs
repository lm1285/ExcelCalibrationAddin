using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Generation
{
    internal sealed class MeasurementGenerationSession
    {
        private readonly Dictionary<string, SharedMeasurementProfile> _profiles =
            new Dictionary<string, SharedMeasurementProfile>(StringComparer.Ordinal);
        private readonly Dictionary<string, SharedMeasurementProfile> _crossItemProfiles =
            new Dictionary<string, SharedMeasurementProfile>(StringComparer.Ordinal);
        private readonly Dictionary<string, List<GeneratedTrendError>> _trendErrors =
            new Dictionary<string, List<GeneratedTrendError>>(StringComparer.Ordinal);
        private readonly Dictionary<string, Dictionary<long, CellMeta>> _cellsBySheet =
            new Dictionary<string, Dictionary<long, CellMeta>>(StringComparer.OrdinalIgnoreCase);

        public MeasurementGenerationSession(WorkbookSnapshot snapshot)
        {
            foreach (var sheet in snapshot?.Sheets ?? new List<SheetSnapshot>())
            {
                _cellsBySheet[sheet.Name ?? string.Empty] = (sheet.Cells ?? new List<CellMeta>())
                    .GroupBy(cell => BuildCellKey(cell.Row, cell.Column))
                    .ToDictionary(group => group.Key, group => group.Last());
            }
        }

        public string BuildKey(MeasurementRule rule, double standardValue, string resolvedUnit)
        {
            var unit = (resolvedUnit ?? rule?.FormatRule?.UnitSuffix ?? string.Empty).Trim().ToUpperInvariant();
            var referenceRange = rule?.FixedReferenceRange;
            var coefficients = rule?.GenerationCoefficientOverride;
            return string.Join("|", new[]
            {
                standardValue.ToString("R", CultureInfo.InvariantCulture),
                unit,
                referenceRange?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                rule?.FixedMpe?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                rule?.FixedNegativeTolerance?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                rule?.FixedPositiveTolerance?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                rule?.ErrorType.ToString() ?? string.Empty,
                rule?.ErrorFormula?.Scale.ToString() ?? string.Empty,
                rule?.RequirementOperator.ToString() ?? string.Empty,
                rule?.MeasurementLowerBound?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                rule?.MeasurementUpperBound?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                coefficients?.NegativeMinimumCoefficient?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                coefficients?.NegativeMaximumCoefficient?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                coefficients?.PositiveMinimumCoefficient?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                coefficients?.PositiveMaximumCoefficient?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty
            });
        }

        public SharedMeasurementProfile Find(string key)
        {
            return key != null && _profiles.TryGetValue(key, out var profile) ? profile : null;
        }

        public string BuildCrossItemKey(MeasurementRule rule, double standardValue, string resolvedUnit)
        {
            var unit = (resolvedUnit ?? rule?.FormatRule?.UnitSuffix ?? string.Empty).Trim().ToUpperInvariant();
            return string.Join("|", new[]
            {
                standardValue.ToString("R", CultureInfo.InvariantCulture),
                unit
            });
        }

        public SharedMeasurementProfile FindCrossItem(string key)
        {
            return key != null && _crossItemProfiles.TryGetValue(key, out var profile) ? profile : null;
        }

        public string BuildTrendKey(MeasurementRule rule, string resolvedUnit)
        {
            // FieldName is the stable template item identity. FieldAlias can
            // be row-specific after recognition and must not split one trend
            // group into independent rules.
            var itemName = string.IsNullOrWhiteSpace(rule?.FieldName)
                ? rule?.FieldAlias
                : rule.FieldName;
            var unit = (resolvedUnit ?? rule?.FormatRule?.UnitSuffix ?? string.Empty).Trim().ToUpperInvariant();
            var formula = rule?.ErrorFormula;
            var section = rule?.TemplateDefinition?.SectionRange;
            return string.Join("|", new[]
            {
                itemName ?? string.Empty,
                unit,
                section?.SheetName ?? string.Empty,
                section?.StartRow.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                section?.EndRow.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                rule?.ErrorType.ToString() ?? string.Empty,
                formula?.Scale.ToString() ?? string.Empty,
                rule?.RequirementOperator.ToString() ?? string.Empty,
                rule?.FixedMpe?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                rule?.FixedNegativeTolerance?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                rule?.FixedPositiveTolerance?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty,
                rule?.FixedReferenceRange?.ToString("R", CultureInfo.InvariantCulture) ?? string.Empty
            });
        }

        public IReadOnlyList<GeneratedTrendError> GetTrendErrors(string key)
        {
            return key != null && _trendErrors.TryGetValue(key, out var errors)
                ? errors.ToList()
                : new List<GeneratedTrendError>();
        }

        public void AddTrendError(string key, double standardValue, double value)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return;
            }

            if (!_trendErrors.TryGetValue(key, out var errors))
            {
                errors = new List<GeneratedTrendError>();
                _trendErrors[key] = errors;
            }

            if (!errors.Any(item =>
                    Math.Abs(item.StandardValue - standardValue) <= 1e-12 &&
                    Math.Abs(item.Value - value) <= 1e-12))
            {
                errors.Add(new GeneratedTrendError(standardValue, value));
            }
        }

        public CellMeta FindCell(string sheetName, int row, int column)
        {
            return _cellsBySheet.TryGetValue(sheetName ?? string.Empty, out var cells) &&
                cells.TryGetValue(BuildCellKey(row, column), out var cell)
                ? cell
                : null;
        }

        public void Store(
            string key,
            string crossItemKey,
            MeasurementGenerationResult result,
            double representativeError)
        {
            if (string.IsNullOrWhiteSpace(key) || result == null)
            {
                return;
            }

            var profile = new SharedMeasurementProfile
            {
                RawValues = result.RawValues.ToList(),
                Direction = result.Direction,
                RepresentativeError = representativeError
            };
            _profiles[key] = profile;
            if (!string.IsNullOrWhiteSpace(crossItemKey))
            {
                _crossItemProfiles[crossItemKey] = profile;
            }
        }

        private static long BuildCellKey(int row, int column)
        {
            return ((long)row << 32) | (uint)column;
        }
    }

    internal sealed class SharedMeasurementProfile
    {
        public List<double> RawValues { get; set; } = new List<double>();
        public int Direction { get; set; }
        public double RepresentativeError { get; set; }
    }

    internal sealed class GeneratedTrendError
    {
        public GeneratedTrendError(double standardValue, double value)
        {
            StandardValue = standardValue;
            Value = value;
        }

        public double StandardValue { get; }
        public double Value { get; }
    }
}
