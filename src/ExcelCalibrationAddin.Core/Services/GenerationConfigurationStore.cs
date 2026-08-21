using System;
using System.IO;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class GenerationConfigurationStore
    {
        private readonly ConfigurationLoader _loader;

        public GenerationConfigurationStore()
            : this(new ConfigurationLoader())
        {
        }

        public GenerationConfigurationStore(ConfigurationLoader loader)
        {
            _loader = loader ?? new ConfigurationLoader();
        }

        public string GetUserConfigurationPath()
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppData, "ExcelCalibrationAddin", "random-generation.json");
        }

        public GenerationConfiguration Load(string baseConfigPath)
        {
            var configuration = _loader.Load(baseConfigPath).Generation ?? new GenerationConfiguration();
            var userConfigurationPath = GetUserConfigurationPath();
            if (File.Exists(userConfigurationPath))
            {
                var content = File.ReadAllText(userConfigurationPath);
                var saved = JsonConvert.DeserializeObject<GenerationConfiguration>(content);
                if (saved != null)
                {
                    configuration = saved;
                }
            }

            return Normalize(configuration);
        }

        public void Save(GenerationConfiguration configuration)
        {
            var normalized = Normalize(configuration);
            var path = GetUserConfigurationPath();
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(path, JsonConvert.SerializeObject(normalized, Formatting.Indented));
        }

        public GenerationConfiguration Clone(GenerationConfiguration configuration)
        {
            var normalized = Normalize(configuration);
            return new GenerationConfiguration
            {
                GenerateShortcutKey = normalized.GenerateShortcutKey,
                DefaultDistribution = normalized.DefaultDistribution,
                ResultCalculationMethod = normalized.ResultCalculationMethod,
                StandardValueReference = normalized.StandardValueReference,
                UseSameDeviationDirection = normalized.UseSameDeviationDirection,
                UseIndependentDeviationControl = normalized.UseIndependentDeviationControl,
                UnifiedErrorMinimumCoefficient = normalized.UnifiedErrorMinimumCoefficient,
                UnifiedErrorMaximumCoefficient = normalized.UnifiedErrorMaximumCoefficient,
                PositiveErrorMinimumCoefficient = normalized.PositiveErrorMinimumCoefficient,
                PositiveErrorMaximumCoefficient = normalized.PositiveErrorMaximumCoefficient,
                NegativeErrorMinimumCoefficient = normalized.NegativeErrorMinimumCoefficient,
                NegativeErrorMaximumCoefficient = normalized.NegativeErrorMaximumCoefficient,
                AbsoluteErrorMinimumCoefficient = normalized.AbsoluteErrorMinimumCoefficient,
                AbsoluteErrorMaximumCoefficient = normalized.AbsoluteErrorMaximumCoefficient,
                MinimumRequirementMinimumCoefficient = normalized.MinimumRequirementMinimumCoefficient,
                MinimumRequirementMaximumCoefficient = normalized.MinimumRequirementMaximumCoefficient,
                UseDecimalPlacesForResolution = normalized.UseDecimalPlacesForResolution,
                MeasurementGroupMinimumFluctuationCoefficient = normalized.MeasurementGroupMinimumFluctuationCoefficient,
                MeasurementGroupMaximumFluctuationCoefficient = normalized.MeasurementGroupMaximumFluctuationCoefficient,
                ResultGroupMinimumFluctuationCoefficient = normalized.ResultGroupMinimumFluctuationCoefficient,
                ResultGroupMaximumFluctuationCoefficient = normalized.ResultGroupMaximumFluctuationCoefficient,
                ResponseTimeThresholdSeconds = normalized.ResponseTimeThresholdSeconds,
                ResponseTimeBelowThresholdMaximumDifferenceSeconds = normalized.ResponseTimeBelowThresholdMaximumDifferenceSeconds,
                ResponseTimeAboveThresholdMaximumDifferenceSeconds = normalized.ResponseTimeAboveThresholdMaximumDifferenceSeconds
            };
        }

        public GenerationConfiguration Normalize(GenerationConfiguration configuration)
        {
            if (configuration == null)
            {
                configuration = new GenerationConfiguration();
            }

            if (string.IsNullOrWhiteSpace(configuration.DefaultDistribution))
            {
                configuration.DefaultDistribution = "Normal";
            }

            configuration.GenerateShortcutKey = NormalizeShortcutKey(configuration.GenerateShortcutKey);

            if (string.Equals(configuration.DefaultDistribution, "TruncatedNormal", StringComparison.OrdinalIgnoreCase))
            {
                configuration.DefaultDistribution = "Normal";
            }

            if (string.IsNullOrWhiteSpace(configuration.ResultCalculationMethod))
            {
                configuration.ResultCalculationMethod = "FormulaBackCalculation";
            }

            if (string.IsNullOrWhiteSpace(configuration.StandardValueReference))
            {
                configuration.StandardValueReference = "RecognizedStandardValueRange";
            }

            configuration.UnifiedErrorMinimumCoefficient = Clamp(configuration.UnifiedErrorMinimumCoefficient, 0.0, 1.0);
            configuration.UnifiedErrorMaximumCoefficient = Clamp(configuration.UnifiedErrorMaximumCoefficient, configuration.UnifiedErrorMinimumCoefficient, 1.0);
            configuration.PositiveErrorMinimumCoefficient = Clamp(configuration.PositiveErrorMinimumCoefficient, 0.0, 1.0);
            configuration.PositiveErrorMaximumCoefficient = Clamp(configuration.PositiveErrorMaximumCoefficient, configuration.PositiveErrorMinimumCoefficient, 1.0);
            configuration.NegativeErrorMinimumCoefficient = Clamp(configuration.NegativeErrorMinimumCoefficient, 0.0, 1.0);
            configuration.NegativeErrorMaximumCoefficient = Clamp(configuration.NegativeErrorMaximumCoefficient, configuration.NegativeErrorMinimumCoefficient, 1.0);
            configuration.AbsoluteErrorMinimumCoefficient = Clamp(configuration.AbsoluteErrorMinimumCoefficient, 0.0, 1.0);
            configuration.AbsoluteErrorMaximumCoefficient = Clamp(configuration.AbsoluteErrorMaximumCoefficient, configuration.AbsoluteErrorMinimumCoefficient, 1.0);
            configuration.MinimumRequirementMinimumCoefficient = Clamp(configuration.MinimumRequirementMinimumCoefficient, 1.0, 10.0);
            configuration.MinimumRequirementMaximumCoefficient = Clamp(configuration.MinimumRequirementMaximumCoefficient, configuration.MinimumRequirementMinimumCoefficient, 10.0);
            configuration.UseDecimalPlacesForResolution = true;
            configuration.MeasurementGroupMinimumFluctuationCoefficient = Clamp(configuration.MeasurementGroupMinimumFluctuationCoefficient, 0.01, 1.0);
            configuration.MeasurementGroupMaximumFluctuationCoefficient = Clamp(configuration.MeasurementGroupMaximumFluctuationCoefficient, 0.0, 1.0);
            configuration.MeasurementGroupMaximumFluctuationCoefficient = Math.Max(
                configuration.MeasurementGroupMinimumFluctuationCoefficient,
                configuration.MeasurementGroupMaximumFluctuationCoefficient);
            configuration.ResultGroupMinimumFluctuationCoefficient = Clamp(configuration.ResultGroupMinimumFluctuationCoefficient, 0.0, 1.0);
            configuration.ResultGroupMaximumFluctuationCoefficient = Clamp(
                configuration.ResultGroupMaximumFluctuationCoefficient,
                configuration.ResultGroupMinimumFluctuationCoefficient,
                1.0);
            configuration.ResponseTimeThresholdSeconds = Clamp(configuration.ResponseTimeThresholdSeconds, 0.01, 100000.0);
            configuration.ResponseTimeBelowThresholdMaximumDifferenceSeconds = Clamp(
                configuration.ResponseTimeBelowThresholdMaximumDifferenceSeconds,
                0.01,
                100000.0);
            configuration.ResponseTimeAboveThresholdMaximumDifferenceSeconds = Clamp(
                configuration.ResponseTimeAboveThresholdMaximumDifferenceSeconds,
                0.01,
                100000.0);
            return configuration;
        }

        private static string NormalizeShortcutKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var parts = value.Trim().Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
            {
                return "F6";
            }

            var useControl = false;
            var useAlt = false;
            var useShift = false;
            for (var index = 0; index < parts.Length - 1; index++)
            {
                var modifier = parts[index].Trim();
                if (string.Equals(modifier, "Ctrl", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(modifier, "Control", StringComparison.OrdinalIgnoreCase))
                {
                    if (useControl)
                    {
                        return "F6";
                    }

                    useControl = true;
                }
                else if (string.Equals(modifier, "Alt", StringComparison.OrdinalIgnoreCase))
                {
                    if (useAlt)
                    {
                        return "F6";
                    }

                    useAlt = true;
                }
                else if (string.Equals(modifier, "Shift", StringComparison.OrdinalIgnoreCase))
                {
                    if (useShift)
                    {
                        return "F6";
                    }

                    useShift = true;
                }
                else
                {
                    return "F6";
                }
            }

            var key = parts[parts.Length - 1].Trim();
            if (string.IsNullOrWhiteSpace(key) || key.IndexOfAny(new[] { '+', ' ' }) >= 0)
            {
                return "F6";
            }

            if (!useControl && !useAlt && !useShift && !IsFunctionKey(key))
            {
                return "F6";
            }

            var modifiers = new System.Collections.Generic.List<string>();
            if (useControl) modifiers.Add("Ctrl");
            if (useAlt) modifiers.Add("Alt");
            if (useShift) modifiers.Add("Shift");
            modifiers.Add(NormalizeKeyName(key));
            return string.Join("+", modifiers);
        }

        private static bool IsFunctionKey(string value)
        {
            var normalized = value.Trim().ToUpperInvariant();
            for (var index = 1; index <= 12; index++)
            {
                if (string.Equals(normalized, "F" + index, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizeKeyName(string value)
        {
            var normalized = value.Trim();
            if (normalized.Length == 1 || IsFunctionKey(normalized))
            {
                return normalized.ToUpperInvariant();
            }

            return normalized;
        }

        private static int Clamp(int value, int minimum, int maximum)
        {
            return Math.Max(minimum, Math.Min(maximum, value));
        }

        private static double Clamp(double value, double minimum, double maximum)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return minimum;
            }

            return Math.Max(minimum, Math.Min(maximum, value));
        }
    }
}
