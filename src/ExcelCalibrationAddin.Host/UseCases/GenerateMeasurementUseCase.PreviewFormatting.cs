using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Generation;
using ExcelCalibrationAddin.Host.Services;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public sealed partial class GenerateMeasurementUseCase
    {
        private static IReadOnlyList<int> ResolveDecimalPlaces(
            MeasurementGenerationSession session,
            MeasurementRule rule,
            IReadOnlyList<CellAddress> measurementCells)
        {
            // Precision is read from the current target cell on every
            // generation. A stale template format is not authoritative;
            // unreadable formats use the documented one-decimal fallback.
            var fallback = 1;
            var interpreter = new NumberFormatInterpreter();
            return measurementCells.Select(address =>
            {
                var cell = session.FindCell(rule?.TargetRange?.SheetName, address.Row, address.Column);
                var formatPlaces = interpreter.Interpret(cell?.NumberFormat).DecimalPlaces;
                return formatPlaces ?? ResolveDisplayedDecimalPlaces(cell?.DisplayText) ?? fallback;
            }).ToList();
        }

        private static int ResolveErrorDecimalPlaces(
            MeasurementGenerationSession session,
            MeasurementRule rule)
        {
            var range = rule?.ErrorSource?.Range;
            var fallback = 1;
            if (range == null)
            {
                return fallback;
            }

            var interpreter = new NumberFormatInterpreter();
            for (var row = range.StartRow; row <= range.EndRow; row++)
            {
                for (var column = range.StartColumn; column <= range.EndColumn; column++)
                {
                    var cell = session.FindCell(range.SheetName, row, column);
                    var formatPlaces = interpreter.Interpret(cell?.NumberFormat).DecimalPlaces;
                    var displayedPlaces = ResolveDisplayedDecimalPlaces(cell?.DisplayText);
                    if (formatPlaces.HasValue || displayedPlaces.HasValue)
                    {
                        return formatPlaces ?? displayedPlaces.Value;
                    }
                }
            }

            return fallback;
        }

        private static string ResolveMeasurementUnit(MeasurementGenerationSession session, MeasurementRule rule)
        {
            // Read the active unit signal from the current workbook snapshot
            // first.  A formula-backed unit can change between generations,
            // so the format captured during recognition is only a diagnostic
            // fallback and never the runtime authority.
            var ranges = new[]
            {
                rule?.MpeSource?.Range,
                rule?.ErrorSource?.Range,
                rule?.StandardValueSource?.Range,
                rule?.TargetRange
            }.Where(range => range != null).ToList();
            foreach (var range in ranges)
            {
                for (var row = range.StartRow; row <= range.EndRow; row++)
                {
                    for (var column = range.StartColumn; column <= range.EndColumn; column++)
                    {
                        var cell = session.FindCell(range.SheetName, row, column);
                        var unit = TemplateUnitParser.Extract(cell?.DisplayText, cell?.Text, cell?.RawValueText, cell?.NumberFormat);
                        if (!string.IsNullOrWhiteSpace(unit))
                        {
                            return unit.Trim();
                        }
                    }
                }
            }

            return (rule?.FormatRule?.UnitSuffix ?? string.Empty).Trim();
        }

        private static int? ResolveDisplayedDecimalPlaces(string text)
        {
            var match = Regex.Match(text ?? string.Empty, @"[-+]?\d+(?:\.(\d+))?");
            return match.Success ? match.Groups[1].Value.Length : (int?)null;
        }

        private static void EnsureValuesMatchResolution(
            MeasurementRule rule,
            IReadOnlyList<double> values,
            IReadOnlyList<int> decimalPlaces)
        {
            for (var index = 0; index < values.Count; index++)
            {
                var places = decimalPlaces[index];
                if (Math.Abs(values[index] - Math.Round(values[index], places)) > 1e-12)
                {
                    throw new InvalidOperationException(
                        $"“{GenerationRuleValidator.ResolveRuleName(rule)}”与同标准值项目的小数分辨力不兼容，无法复用同一批测量值。");
                }
            }
        }

        private static void EnsureSharedDirectionIsCompatible(MeasurementRule rule, int direction)
        {
            if ((rule?.PositiveDirectionOnly == true && direction < 0) ||
                (rule?.NegativeDirectionOnly == true && direction > 0))
            {
                throw new InvalidOperationException(
                    $"“{GenerationRuleValidator.ResolveRuleName(rule)}”的强制误差方向与同标准值项目冲突。");
            }
        }

        private static void EnsureValuesMatchIndependentInterval(
            MeasurementRule rule,
            IReadOnlyList<double> values)
        {
            if (rule?.MeasurementLowerBound.HasValue == true &&
                values.Any(value => value < rule.MeasurementLowerBound.Value - 1e-12))
            {
                throw new InvalidOperationException(
                    $"“{GenerationRuleValidator.ResolveRuleName(rule)}”与同标准值项目的独立测量下限冲突。");
            }

            if (rule?.MeasurementUpperBound.HasValue == true &&
                values.Any(value => value > rule.MeasurementUpperBound.Value + 1e-12))
            {
                throw new InvalidOperationException(
                    $"“{GenerationRuleValidator.ResolveRuleName(rule)}”与同标准值项目的独立测量上限冲突。");
            }
        }

        private static MeasurementGenerationResult BuildGenerationResult(
            IReadOnlyList<double> rawValues,
            IReadOnlyList<int> decimalPlaces,
            int direction)
        {
            return new MeasurementGenerationResult
            {
                RawValues = rawValues.ToList(),
                DisplayValues = rawValues.Select((value, index) =>
                    Math.Round(value, decimalPlaces[index]).ToString($"F{decimalPlaces[index]}", CultureInfo.InvariantCulture)).ToList(),
                Direction = direction
            };
        }

        private static double CalculateRoundedFormulaError(
            MeasurementRule rule,
            double standardValue,
            IReadOnlyList<double> writtenValues,
            int errorDecimalPlaces)
        {
            return Math.Round(
                CalculateFormulaError(rule, standardValue, writtenValues),
                Math.Max(0, Math.Min(15, errorDecimalPlaces)));
        }
    }
}
