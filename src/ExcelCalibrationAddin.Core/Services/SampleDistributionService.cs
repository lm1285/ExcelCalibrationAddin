using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Core.Models;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class SampleDistributionService
    {
        private readonly Random _random;
        public SampleDistributionService(Random random = null) { _random = random ?? new Random(); }

        public bool TryGenerate(IEnumerable<SampleDataPoint> points, string calibrationItemName, double? standardValue, int count, out List<double> values, out int decimalPlaces)
        {
            values = new List<double>(); decimalPlaces = 0;
            var candidates = (points ?? Enumerable.Empty<SampleDataPoint>()).Where(point => point != null && (string.IsNullOrWhiteSpace(calibrationItemName) || string.Equals(point.CalibrationItemName, calibrationItemName, StringComparison.OrdinalIgnoreCase)) && point.MeasurementValues != null && point.MeasurementValues.Count >= 3 && (!standardValue.HasValue || !point.StandardValue.HasValue || Math.Abs(point.StandardValue.Value - standardValue.Value) <= Math.Max(1e-9, Math.Abs(standardValue.Value) * 1e-9))).ToList();
            var source = candidates.SelectMany(point => point.MeasurementValues).Where(value => !double.IsNaN(value) && !double.IsInfinity(value)).ToList();
            if (source.Count < 3 || count <= 0) return false;
            decimalPlaces = candidates.Select(point => point.DecimalPlaces).DefaultIfEmpty(0).Max();
            var mean = source.Average();
            var variance = source.Select(value => (value - mean) * (value - mean)).Sum() / Math.Max(1, source.Count - 1);
            var deviation = Math.Sqrt(variance);
            var min = source.Min(); var max = source.Max(); var span = Math.Max(max - min, Math.Pow(10, -decimalPlaces));
            var lower = min - span * 0.1; var upper = max + span * 0.1;
            for (var i = 0; i < count; i++)
            {
                var sampled = deviation <= 1e-12 ? mean : mean + deviation * NextGaussian();
                sampled = Math.Max(lower, Math.Min(upper, sampled));
                values.Add(Math.Round(sampled, decimalPlaces, MidpointRounding.AwayFromZero));
            }
            return true;
        }

        private double NextGaussian()
        {
            var u1 = Math.Max(double.Epsilon, _random.NextDouble());
            var u2 = _random.NextDouble();
            return Math.Sqrt(-2d * Math.Log(u1)) * Math.Cos(2d * Math.PI * u2);
        }
    }
}
