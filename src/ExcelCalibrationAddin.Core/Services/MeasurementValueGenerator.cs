using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed partial class MeasurementValueGenerator
    {
        private readonly Random _random;
        private readonly GenerationConfiguration _configuration;

        public MeasurementValueGenerator()
            : this(new GenerationConfiguration(), CreateRandom())
        {
        }

        public MeasurementValueGenerator(Random random)
            : this(new GenerationConfiguration(), random)
        {
        }

        public MeasurementValueGenerator(GenerationConfiguration configuration)
            : this(configuration, CreateRandom())
        {
        }

        public MeasurementValueGenerator(GenerationConfiguration configuration, Random random)
        {
            _configuration = new GenerationConfigurationStore().Normalize(configuration);
            _random = random ?? CreateRandom();
        }

        public MeasurementGenerationResult Generate(MeasurementGenerationInput input)
        {
            ValidateInput(input);
            var bounds = ResolveBounds(input);
            var direction = ResolveDirection(input, bounds.lower, bounds.upper);
            var errors = GenerateErrorGroup(input, bounds.lower, bounds.upper, direction);
            if (errors.Any(error => Math.Abs(error) <= 1e-12))
            {
                throw new InvalidOperationException(
                    "Unable to generate a non-zero calibration error at the configured measurement resolution.");
            }
            var rawValues = errors.Select(error => input.StandardValue + error).ToList();
            var displayValues = rawValues
                .Select((value, index) =>
                {
                    var decimalPlaces = ResolveDecimalPlaces(input, index);
                    return Math.Round(value, decimalPlaces).ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture);
                })
                .ToList();
            if (input.ValueCount > 1 &&
                rawValues
                .Select((value, index) => Math.Round(value, ResolveDecimalPlaces(input, index)))
                .Select(value => Math.Round(value, 15))
                .Distinct()
                .Count() <= 1)
            {
                throw new InvalidOperationException(
                    "Unable to generate visibly different repeated measurements within the configured error interval and measurement resolution.");
            }

            return new MeasurementGenerationResult
            {
                RawValues = rawValues,
                DisplayValues = displayValues,
                LowerBound = bounds.lower,
                UpperBound = bounds.upper,
                Direction = direction
            };
        }

        private static Random CreateRandom()
        {
            var seedBytes = new byte[sizeof(int)];
            using (var generator = RandomNumberGenerator.Create())
            {
                generator.GetBytes(seedBytes);
            }

            return new Random(BitConverter.ToInt32(seedBytes, 0));
        }

    }
}
