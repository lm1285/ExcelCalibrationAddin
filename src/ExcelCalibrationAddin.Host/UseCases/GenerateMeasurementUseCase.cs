using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Generation;
using ExcelCalibrationAddin.Host.Recognition;
using ExcelCalibrationAddin.Host.Services;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public interface IWorkbookWriter
    {
        void Write(CellRange range, IReadOnlyList<string> values);
        void Write(CellRange range, IReadOnlyList<CellAddress> writableCells, IReadOnlyList<string> values);
    }

    public sealed partial class GenerateMeasurementUseCase
    {
        private static readonly Regex NumberRegex = new Regex(@"[-+]?\d+(\.\d+)?([eE][-+]?\d+)?", RegexOptions.Compiled);

        private readonly IWorkbookWriter _writer;
        private readonly IWorkbookSnapshotProvider _snapshotProvider;
        private readonly MeasurementRuleParameterResolver _parameterResolver;
        private readonly MeasurementRuleStructureAnalyzer _structureAnalyzer;
        private readonly Func<GenerationConfiguration, MeasurementValueGenerator> _generatorFactory;
        private readonly GenerationConfigurationStore _configurationStore;
        private readonly MeasurementSeriesGenerator _seriesGenerator;
        private GenerationConfiguration _generationConfiguration;
        private readonly Random _random = new Random();

        public GenerateMeasurementUseCase(
            Func<GenerationConfiguration, MeasurementValueGenerator> generatorFactory,
            GenerationConfiguration generationConfiguration,
            IWorkbookWriter writer,
            IWorkbookSnapshotProvider snapshotProvider,
            MeasurementRuleParameterResolver parameterResolver)
        {
            _generatorFactory = generatorFactory ?? throw new ArgumentNullException(nameof(generatorFactory));
            _configurationStore = new GenerationConfigurationStore();
            _seriesGenerator = new MeasurementSeriesGenerator(_random);
            _generationConfiguration = _configurationStore.Clone(generationConfiguration);
            _writer = writer;
            _snapshotProvider = snapshotProvider;
            _parameterResolver = parameterResolver;
            _structureAnalyzer = new MeasurementRuleStructureAnalyzer();
        }

        public void SetGenerationConfiguration(GenerationConfiguration configuration)
        {
            _generationConfiguration = _configurationStore.Clone(configuration);
        }

        public GenerationWriteResult Write(IEnumerable<MeasurementRule> rules)
        {
            var previews = Preview(rules);
            return WritePreviews(previews);
        }

        public GenerationWriteResult WriteResolved(IEnumerable<MeasurementRule> rules)
        {
            var previews = PreviewResolved(rules);
            return WritePreviews(previews);
        }

        public GenerationWriteResult WritePreResolved(IEnumerable<MeasurementRule> rules)
        {
            var previews = PreviewPreResolved(rules);
            return WritePreviews(previews);
        }

        private GenerationWriteResult WritePreviews(IReadOnlyList<RulePreview> previews)
        {
            var skippedWarnings = new List<string>();
            foreach (var preview in previews ?? Array.Empty<RulePreview>())
            {
                if (preview == null || preview.DisplayValues == null || preview.DisplayValues.Count == 0)
                {
                    if (preview == null || preview.WarningMessages == null || preview.WarningMessages.Count == 0)
                    {
                        var ruleName = GenerationRuleValidator.ResolveRuleName(preview?.Rule);
                        skippedWarnings.Add($"“{ruleName}”没有生成可写入的随机数，已跳过写入。请检查标准值是否为空。");
                    }
                    continue;
                }

                _writer.Write(preview.TargetRange, preview.WritableCells, preview.DisplayValues);
            }

            var result = GenerationWriteResult.FromPreviews(previews);
            result.WarningMessages = result.WarningMessages
                .Concat(skippedWarnings)
                .Distinct(StringComparer.Ordinal)
                .ToList();
            return result;
        }

    }

    public sealed class RulePreview
    {
        public MeasurementRule Rule { get; set; }
        public CellRange TargetRange { get; set; }
        public IReadOnlyList<string> DisplayValues { get; set; } = new List<string>();
        public IReadOnlyList<double> RawValues { get; set; } = new List<double>();
        public IReadOnlyList<CellAddress> WritableCells { get; set; } = new List<CellAddress>();
        public IReadOnlyList<string> WarningMessages { get; set; } = new List<string>();
    }

    public sealed class GenerationWriteResult
    {
        public IReadOnlyList<string> WarningMessages { get; set; } = new List<string>();

        public static GenerationWriteResult FromPreviews(IEnumerable<RulePreview> previews)
        {
            var warnings = (previews ?? Enumerable.Empty<RulePreview>())
                .SelectMany(preview => preview?.WarningMessages ?? new List<string>())
                .Where(message => !string.IsNullOrWhiteSpace(message))
                .Distinct(StringComparer.Ordinal)
                .ToList();

            return new GenerationWriteResult
            {
                WarningMessages = warnings
            };
        }
    }
}
