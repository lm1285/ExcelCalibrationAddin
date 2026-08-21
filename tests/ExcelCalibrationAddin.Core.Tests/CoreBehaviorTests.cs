using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Recognition;
using ExcelCalibrationAddin.Host.Services;
using ExcelCalibrationAddin.Host.Generation;
using ExcelCalibrationAddin.Host.Templates;
using ExcelCalibrationAddin.Host.UseCases;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Tests
{
    [TestClass]
    public sealed partial class CoreBehaviorTests
    {
        [TestMethod]
        public void AutomaticMatchSkipsUnrelatedWorkbookSheets()
        {
            var candidates = AutomaticMatchSheetSelector.Select(
                new[] { "台账", "汇总" },
                new[] { "原始记录" },
                hasEnabledTemplates: true,
                hasTemplateWithoutSheetMetadata: false);

            Assert.AreEqual(0, candidates.Count);
        }

        [TestMethod]
        public void AutomaticMatchKeepsLegacyTemplatesWithoutSheetMetadataCompatible()
        {
            var candidates = AutomaticMatchSheetSelector.Select(
                new[] { "台账", "汇总" },
                Array.Empty<string>(),
                hasEnabledTemplates: true,
                hasTemplateWithoutSheetMetadata: true);

            CollectionAssert.AreEqual(new[] { "台账", "汇总" }, candidates.ToArray());
        }

        [TestMethod]
        public void ConfigurationLoaderRestoresRequiredBlankPaths()
        {
            var directory = Path.Combine(Path.GetTempPath(), "ExcelCalibrationAddin.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var configPath = Path.Combine(directory, "appsettings.json");
            File.WriteAllText(configPath, "{\"Backend\":{\"BaseUrl\":\"  \",\"TemplateApiPrefix\":\"  \"},\"Cache\":{\"SqliteFile\":\"  \"}}");

            var configuration = new ConfigurationLoader().Load(configPath);

            Assert.AreEqual("http://localhost:3002", configuration.Backend.BaseUrl);
            Assert.AreEqual("/api/excel-templates", configuration.Backend.TemplateApiPrefix);
            StringAssert.EndsWith(configuration.Cache.SqliteFile, Path.Combine("ExcelCalibrationAddin", "cache.db"));
        }

        [TestMethod]
        public void TemplateCacheSupportsSemicolonInDatabasePath()
        {
            var path = Path.Combine(
                Path.GetTempPath(),
                "ExcelCalibrationAddin.Tests",
                Guid.NewGuid().ToString("N"),
                "cache;review.sqlite");
            var repository = new LocalTemplateRuleCacheRepository(path);

            repository.Initialize();

            Assert.IsTrue(File.Exists(path), "SQLite should use the complete configured path.");
        }

        [TestMethod]
        public void AutomaticMatchSkipsWhenTemplateLibraryIsEmpty()
        {
            var candidates = AutomaticMatchSheetSelector.Select(
                new[] { "原始记录" },
                new[] { "原始记录" },
                hasEnabledTemplates: false,
                hasTemplateWithoutSheetMetadata: false);

            Assert.AreEqual(0, candidates.Count);
        }

        [TestMethod]
        public void FingerprintIgnoresMeasurementAndStandardValues()
        {
            var first = BuildSnapshot(10, 11);
            var second = BuildSnapshot(30, 99);

            var builder = new TemplateFingerprintBuilder();

            Assert.AreEqual(builder.Build(first).ExactFingerprint, builder.Build(second).ExactFingerprint);
        }

        [TestMethod]
        public void FingerprintIgnoresFormulaAndItsCalculatedResult()
        {
            var first = BuildSnapshot(10, 11);
            var second = BuildSnapshot(10, 11);
            first.Sheets[0].Cells[4].Formula = "=C5-A5";
            first.Sheets[0].Cells[4].Text = "1";
            second.Sheets[0].Cells[4].Formula = "=C5-B5";
            second.Sheets[0].Cells[4].Text = "F";

            var builder = new TemplateFingerprintBuilder();

            Assert.AreEqual(builder.Build(first).ExactFingerprint, builder.Build(second).ExactFingerprint);
        }

        [TestMethod]
        public void FingerprintMatchesWhenWorkbookStructureIsShiftedAsWhole()
        {
            var baseline = BuildSnapshot(10, 11);
            var shifted = BuildSnapshot(10, 11);
            var sheet = shifted.Sheets[0];
            sheet.UsedRangeShape = "B2:E9";
            foreach (var cell in sheet.Cells)
            {
                cell.Row++;
                cell.Column++;
            }

            foreach (var header in sheet.Headers)
            {
                header.Column++;
            }

            var builder = new TemplateFingerprintBuilder();

            Assert.AreEqual(builder.Build(baseline).ExactFingerprint, builder.Build(shifted).ExactFingerprint);
        }

        [TestMethod]
        public void FingerprintIgnoresEmptyDefaultFormattedScanPadding()
        {
            var baseline = BuildSnapshot(10, 11);
            var padded = BuildSnapshot(10, 11);
            foreach (var snapshot in new[] { baseline, padded })
            {
                foreach (var cell in snapshot.Sheets[0].Cells)
                {
                    cell.Row++;
                    cell.Column++;
                }

                foreach (var header in snapshot.Sheets[0].Headers)
                {
                    header.Column++;
                }
            }

            padded.Sheets[0].Cells.Add(new CellMeta
            {
                Row = 1,
                Column = 1,
                NumberFormat = "General"
            });

            var builder = new TemplateFingerprintBuilder();

            Assert.AreEqual(builder.Build(baseline).ExactFingerprint, builder.Build(padded).ExactFingerprint);
        }

        [TestMethod]
        public void WritableCellResolverIncludesFormulaCells()
        {
            var snapshot = BuildSnapshot(10, 11);
            snapshot.Sheets[0].Cells[3].Formula = "=A5";

            var resolution = WritableCellResolver.Resolve(snapshot, Range("C5:C5"));

            Assert.AreEqual(1, resolution.Cells.Count);
            Assert.AreEqual(5, resolution.Cells[0].Row);
            Assert.AreEqual(3, resolution.Cells[0].Column);
        }

        [TestMethod]
        public void WritableCellResolverAcceptsOverlappingSnapshotCells()
        {
            var snapshot = BuildSnapshot(10, 11);
            snapshot.Sheets[0].Cells.Add(new CellMeta
            {
                Row = 5,
                Column = 3,
                NumberFormat = "0.0"
            });

            var resolution = WritableCellResolver.Resolve(snapshot, Range("C5:C5"));

            Assert.AreEqual(1, resolution.Cells.Count);
            Assert.AreEqual(5, resolution.Cells[0].Row);
            Assert.AreEqual(3, resolution.Cells[0].Column);
        }

        [TestMethod]
        public void WritableRowDetectionSkipsMergedHeaderCoveredRows()
        {
            var headerMerge = new CellRange
            {
                SheetName = "Sheet1",
                StartRow = 36,
                EndRow = 37,
                StartColumn = 6,
                EndColumn = 8
            };
            var dataMerge = new CellRange
            {
                SheetName = "Sheet1",
                StartRow = 38,
                EndRow = 38,
                StartColumn = 6,
                EndColumn = 8
            };
            var sheet = new SheetSnapshot
            {
                Name = "Sheet1",
                Cells = new List<CellMeta>
                {
                    new CellMeta { Row = 36, Column = 6, Text = "1", IsMerged = true, MergeRange = headerMerge },
                    new CellMeta { Row = 37, Column = 6, IsMerged = true, MergeRange = headerMerge },
                    new CellMeta { Row = 37, Column = 7, IsMerged = true, MergeRange = headerMerge },
                    new CellMeta { Row = 37, Column = 8, IsMerged = true, MergeRange = headerMerge },
                    new CellMeta { Row = 38, Column = 6, IsMerged = true, MergeRange = dataMerge },
                    new CellMeta { Row = 38, Column = 7, IsMerged = true, MergeRange = dataMerge },
                    new CellMeta { Row = 38, Column = 8, IsMerged = true, MergeRange = dataMerge }
                }
            };

            Assert.IsFalse(SheetRowContentAnalyzer.HasDataInRangeRow(sheet, 37, 6, 8));
            Assert.IsFalse(SheetRowContentAnalyzer.HasNumericDataInRangeRow(sheet, 37, 6, 8));
            Assert.AreEqual(0, SheetRowContentAnalyzer.CountWritableTemplateCellsInRow(sheet, 37, 6, 8));
            Assert.AreEqual(1, SheetRowContentAnalyzer.CountWritableTemplateCellsInRow(sheet, 38, 6, 8));
        }

        [TestMethod]
        public void FingerprintIgnoresStandardValueWithUnit()
        {
            var first = BuildSnapshot(10, 11);
            var second = BuildSnapshot(10, 11);
            first.Sheets[0].Cells[2].Text = "10 ppm";
            second.Sheets[0].Cells[2].Text = "20 ppm";

            var builder = new TemplateFingerprintBuilder();

            Assert.AreEqual(builder.Build(first).ExactFingerprint, builder.Build(second).ExactFingerprint);
        }

        [TestMethod]
        public void FingerprintIgnoresStandardValueIncludedInHeaderPath()
        {
            var first = BuildSnapshot(10, 11);
            var second = BuildSnapshot(20, 11);
            first.Sheets[0].Headers[0].Levels.Add("10 ppm");
            second.Sheets[0].Headers[0].Levels.Add("20 ppm");

            var builder = new TemplateFingerprintBuilder();

            Assert.AreEqual(builder.Build(first).ExactFingerprint, builder.Build(second).ExactFingerprint);
            CollectionAssert.AreEqual(
                builder.Build(first).HeaderTexts.ToArray(),
                builder.Build(second).HeaderTexts.ToArray());
        }

        [TestMethod]
        public void FingerprintChangesWhenHeaderChanges()
        {
            var first = BuildSnapshot(10, 11);
            var second = BuildSnapshot(10, 11);
            second.Sheets[0].Cells[1].Text = "测量列";

            var builder = new TemplateFingerprintBuilder();

            Assert.AreNotEqual(builder.Build(first).ExactFingerprint, builder.Build(second).ExactFingerprint);
        }

        [TestMethod]
        public void TemplateDefinitionCapturesHeadersSplitRequirementUnitsAndFormulaBranches()
        {
            var snapshot = BuildTemplateDefinitionSnapshot();
            var mapping = BuildTemplateDefinitionMapping();

            var definition = new TemplateFieldDefinitionBuilder(new NumberFormatInterpreter())
                .Build(snapshot.Sheets[0], mapping);

            Assert.IsNotNull(definition);
            Assert.IsTrue(definition.Headers.Any(header => header.Text.Contains("\u793A\u503C\u8BEF\u5DEE")));
            Assert.IsTrue(definition.Headers.Any(header => header.Text.Contains("\u6280\u672F\u8981\u6C42")));

            var requirement = definition.Regions.Single(region =>
                region.Role == TemplateRegionRole.TechnicalRequirement);
            Assert.AreEqual(7, requirement.OperatorRange.StartColumn);
            Assert.AreEqual(8, requirement.ValueRange.StartColumn);
            Assert.AreEqual("ppm", requirement.Unit);
            CollectionAssert.Contains(requirement.Units, "ppm");
            CollectionAssert.Contains(requirement.Units, "%FS");
            Assert.AreEqual(2, requirement.Formula.Branches.Count);
            Assert.AreEqual("ppm", requirement.Formula.Branches[0].Unit);
            Assert.AreEqual("%FS", requirement.Formula.Branches[1].Unit);
            Assert.AreEqual(TechnicalRequirementOperator.PlusMinus, requirement.Formula.Branches[0].Operator);
            Assert.AreEqual(TechnicalRequirementOperator.LessThanOrEqual, requirement.Formula.Branches[1].Operator);
            Assert.IsTrue(requirement.Formula.IsFullyParsed);
            Assert.AreEqual(1, requirement.RequirementValues.Count);
            Assert.AreEqual(TechnicalRequirementOperator.PlusMinus, requirement.RequirementValues[0].Operator);
            Assert.AreEqual(2d, requirement.RequirementValues[0].Value.Value, 1e-12);
            Assert.AreEqual("ppm", requirement.RequirementValues[0].Unit);

            var error = definition.Regions.Single(region => region.Role == TemplateRegionRole.ErrorValue);
            Assert.IsTrue(error.Formula.References.Any(reference =>
                reference.Role == TemplateRegionRole.StandardValue));
            Assert.IsTrue(error.Formula.References.Any(reference =>
                reference.Role == TemplateRegionRole.AverageValue));
            Assert.IsTrue(error.Formula.References.Any(reference =>
                reference.Role == TemplateRegionRole.RangeValue));

            var standard = definition.Regions.Single(region => region.Role == TemplateRegionRole.StandardValue);
            Assert.IsNotNull(standard.Formula);
            Assert.AreEqual("=RC[1]*0.5", standard.Formula.FormulaR1C1);

            var range = definition.Regions.Single(region => region.Role == TemplateRegionRole.RangeValue);
            Assert.AreEqual("ppm", range.Unit);
            Assert.IsNotNull(range.Formula);
        }

        [TestMethod]
        public void TemplateDefinitionMatcherRejectsChangedErrorFormula()
        {
            var snapshot = BuildTemplateDefinitionSnapshot();
            var mapping = BuildTemplateDefinitionMapping();
            var builder = new TemplateFieldDefinitionBuilder(new NumberFormatInterpreter());
            var saved = builder.Build(snapshot.Sheets[0], mapping);
            var current = TemplateDefinitionCloner.Clone(saved);
            current.Regions.Single(region => region.Role == TemplateRegionRole.ErrorValue)
                .Formula.FormulaR1C1 = "=RC[-1]-RC[-5]";

            Assert.IsFalse(TemplateFieldDefinitionMatcher.IsCompatible(saved, current));
            Assert.IsTrue(TemplateFieldDefinitionMatcher.IsCompatible(saved, TemplateDefinitionCloner.Clone(saved)));
        }

        [TestMethod]
        public void TemplateDefinitionCombinesSplitOperatorWithFormattedValue()
        {
            var sheet = new SheetSnapshot
            {
                Name = "Sheet1",
                Cells = new List<CellMeta>
                {
                    new CellMeta { Row = 1, Column = 1, Text = "\u793A\u503C\u8BEF\u5DEE" },
                    new CellMeta { Row = 2, Column = 3, Text = "\u6D4B\u91CF\u503C" },
                    new CellMeta { Row = 2, Column = 4, Text = "\u7B26\u53F7" },
                    new CellMeta { Row = 2, Column = 5, Text = "\u6280\u672F\u8981\u6C42(%FS)" },
                    new CellMeta { Row = 3, Column = 3, NumberFormat = "0.0" },
                    new CellMeta { Row = 3, Column = 4, Text = "\u2264", DisplayText = "\u2264" },
                    new CellMeta { Row = 3, Column = 5, Text = "3", DisplayText = "3.0", RawValueText = "3", NumberFormat = "0.0\"%FS\"" }
                }
            };
            var mapping = new TemplateRegionMapping
            {
                ProjectName = "\u793A\u503C\u8BEF\u5DEE",
                SectionRange = new CellRange { SheetName = "Sheet1", StartRow = 1, EndRow = 3, StartColumn = 1, EndColumn = 5 },
                MeasurementValueRange = RangeAt(3, 3),
                TechnicalRequirementRange = new CellRange { SheetName = "Sheet1", StartRow = 3, EndRow = 3, StartColumn = 4, EndColumn = 5 }
            };

            var requirement = new TemplateFieldDefinitionBuilder(new NumberFormatInterpreter())
                .Build(sheet, mapping)
                .Regions.Single(region => region.Role == TemplateRegionRole.TechnicalRequirement);

            Assert.AreEqual(4, requirement.OperatorRange.StartColumn);
            Assert.AreEqual(5, requirement.ValueRange.StartColumn);
            Assert.AreEqual(1, requirement.RequirementValues.Count);
            Assert.AreEqual(TechnicalRequirementOperator.LessThanOrEqual, requirement.RequirementValues[0].Operator);
            Assert.AreEqual(3d, requirement.RequirementValues[0].Value.Value, 1e-12);
            Assert.AreEqual("%FS", requirement.RequirementValues[0].Unit);
        }

        [TestMethod]
        public void ParameterResolverRestoresDynamicSourcesFromTemplateDefinition()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 2, Column = 9, Text = "100", RawValueText = "100" },
                            new CellMeta { Row = 5, Column = 1, Text = "20", RawValueText = "20" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "restored sources",
                TargetRange = RangeAt(5, 3),
                ErrorType = ErrorType.Referenced,
                ErrorFormula = new ErrorFormulaInfo { Scale = ErrorFormulaScale.RelativeToReferenceRange },
                TemplateDefinition = new TemplateFieldDefinition
                {
                    Regions = new List<TemplateRegionDefinition>
                    {
                        new TemplateRegionDefinition { Role = TemplateRegionRole.StandardValue, Range = RangeAt(5, 1) },
                        new TemplateRegionDefinition { Role = TemplateRegionRole.RangeValue, Range = RangeAt(2, 9) }
                    }
                }
            };

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(20d, resolved.FixedStandardValue.Value, 1e-12);
            Assert.AreEqual(100d, resolved.FixedReferenceRange.Value, 1e-12);
            Assert.AreEqual(1, resolved.StandardValueSource.Range.StartColumn);
            Assert.AreEqual(9, resolved.RangeSource.Range.StartColumn);
        }

        [TestMethod]
        public void RelativeGenerationUsesAbsoluteToleranceForNegativeStandardValue()
        {
            var generator = new MeasurementValueGenerator(new GenerationConfiguration(), new Random(7));
            var result = generator.Generate(new MeasurementGenerationInput
            {
                StandardValue = -10,
                Mpe = 0.1,
                ErrorType = ErrorType.Relative,
                ValueCount = 20,
                DecimalPlaces = 3
            });

            Assert.IsTrue(result.RawValues.All(value => value >= -11 && value <= -9));
        }

        [DataTestMethod]
        [DataRow("±0.5", TechnicalRequirementOperator.PlusMinus)]
        [DataRow("≤0.5", TechnicalRequirementOperator.LessThanOrEqual)]
        [DataRow(">=0.5", TechnicalRequirementOperator.GreaterThanOrEqual)]
        public void RequirementParserRecognizesOperators(string text, TechnicalRequirementOperator expected)
        {
            var result = RequirementTextParser.Parse(new CellMeta { Text = text });

            Assert.AreEqual(expected, result.Operator);
        }

        [TestMethod]
        public void RequirementParserReadsOperatorFromCustomNumberFormat()
        {
            var result = RequirementTextParser.Parse(new CellMeta
            {
                RawValueText = "3",
                DisplayText = "3.0",
                NumberFormat = "\"≤\"0.0\"%FS\""
            });

            Assert.AreEqual(TechnicalRequirementOperator.LessThanOrEqual, result.Operator);
        }

        [TestMethod]
        public void FingerprintChangesWhenTechnicalRequirementChanges()
        {
            var first = BuildSnapshot(10, 11);
            var second = BuildSnapshot(10, 11);
            first.Sheets[0].Cells[5].Text = "±0.5";
            second.Sheets[0].Cells[5].Text = "±1.0";

            var builder = new TemplateFingerprintBuilder();

            Assert.AreNotEqual(builder.Build(first).ExactFingerprint, builder.Build(second).ExactFingerprint);
        }

        [TestMethod]
        public void GreaterThanRequirementGeneratesMagnitudeAboveLimit()
        {
            var generator = new MeasurementValueGenerator(new GenerationConfiguration(), new Random(17));
            var result = generator.Generate(new MeasurementGenerationInput
            {
                StandardValue = 10,
                Mpe = 0.5,
                ErrorType = ErrorType.Absolute,
                RequirementOperator = TechnicalRequirementOperator.GreaterThan,
                ValueCount = 3,
                DecimalPlaces = 3
            });

            Assert.IsTrue(result.RawValues.All(value => Math.Abs(value - 10) > 0.5));
            Assert.IsTrue(result.RawValues.All(value => Math.Abs(value - 10) <= 0.65));
        }

        [TestMethod]
        public void DefaultGenerationConfigurationUsesRequestedIntervals()
        {
            var configuration = new GenerationConfiguration();

            Assert.AreEqual(0.2, configuration.PositiveErrorMinimumCoefficient, 1e-12);
            Assert.AreEqual(0.8, configuration.PositiveErrorMaximumCoefficient, 1e-12);
            Assert.AreEqual(0.2, configuration.ResultGroupMaximumFluctuationCoefficient, 1e-12);
            Assert.AreEqual(0.01, configuration.MeasurementGroupMinimumFluctuationCoefficient, 1e-12);
            Assert.AreEqual(0.06, configuration.MeasurementGroupMaximumFluctuationCoefficient, 1e-12);
            Assert.AreEqual(180, configuration.ResponseTimeThresholdSeconds, 1e-12);
            Assert.AreEqual(5, configuration.ResponseTimeBelowThresholdMaximumDifferenceSeconds, 1e-12);
            Assert.AreEqual(7, configuration.ResponseTimeAboveThresholdMaximumDifferenceSeconds, 1e-12);
        }

        [TestMethod]
        public void NormalDistributionDoesNotOverconcentrateAtRoundedMidpoint()
        {
            var configuration = new GenerationConfiguration
            {
                DefaultDistribution = "Normal",
                UseDecimalPlacesForResolution = true,
                PositiveErrorMinimumCoefficient = 0.2,
                PositiveErrorMaximumCoefficient = 0.8
            };
            var generator = new MeasurementValueGenerator(configuration, new Random(20260719));
            const int sampleCount = 4000;
            var midpointCount = 0;

            for (var index = 0; index < sampleCount; index++)
            {
                var result = generator.Generate(new MeasurementGenerationInput
                {
                    StandardValue = 100,
                    Mpe = 10,
                    ErrorType = ErrorType.Absolute,
                    DistributionMode = DistributionMode.Normal,
                    ValueCount = 1,
                    DecimalPlaces = 0,
                    ForcePositiveDirection = true
                });
                var error = result.RawValues.Single() - 100;
                Assert.IsTrue(error >= 2 && error <= 8);
                if (Math.Abs(error - 5) <= 1e-12)
                {
                    midpointCount++;
                }
            }

            Assert.IsTrue(midpointCount < sampleCount * 0.33);
        }

        [TestMethod]
        public void LegacyGenerationConfigurationMigratesGroupFluctuationMaximum()
        {
            var legacy = JsonConvert.DeserializeObject<GenerationConfiguration>(
                "{\"MeasurementGroupFluctuationCoefficient\":0.08,\"CrossStandardValueFluctuationCoefficient\":0.15}");
            var normalized = new GenerationConfigurationStore().Normalize(legacy);
            var saved = JsonConvert.SerializeObject(normalized);

            Assert.AreEqual(0.01, normalized.MeasurementGroupMinimumFluctuationCoefficient, 1e-12);
            Assert.AreEqual(0.08, normalized.MeasurementGroupMaximumFluctuationCoefficient, 1e-12);
            Assert.AreEqual(0.15, normalized.ResultGroupMaximumFluctuationCoefficient, 1e-12);
            Assert.IsFalse(saved.Contains("MeasurementGroupFluctuationCoefficient"));
            Assert.IsFalse(saved.Contains("CrossStandardValueFluctuationCoefficient"));
            Assert.IsFalse(saved.Contains("ResultGroupFluctuationCoefficient"));
            Assert.IsFalse(saved.Contains("UseDecimalPlacesForResolution"));
            StringAssert.Contains(saved, "MeasurementGroupMinimumFluctuationCoefficient");
        }

        [TestMethod]
        public void AbsoluteErrorUsesIndependentPositiveAndNegativeIntervals()
        {
            var configuration = new GenerationConfiguration
            {
                PositiveErrorMinimumCoefficient = 0.2,
                PositiveErrorMaximumCoefficient = 0.25,
                NegativeErrorMinimumCoefficient = 0.7,
                NegativeErrorMaximumCoefficient = 0.8
            };
            var positive = new MeasurementValueGenerator(configuration, new Random(31)).Generate(new MeasurementGenerationInput
            {
                StandardValue = 100,
                Mpe = 10,
                ErrorType = ErrorType.Absolute,
                ForcePositiveDirection = true,
                ValueCount = 3,
                DecimalPlaces = 2
            });
            var negative = new MeasurementValueGenerator(configuration, new Random(31)).Generate(new MeasurementGenerationInput
            {
                StandardValue = 100,
                Mpe = 10,
                ErrorType = ErrorType.Absolute,
                ForceNegativeDirection = true,
                ValueCount = 3,
                DecimalPlaces = 2
            });

            Assert.IsTrue(positive.RawValues.All(value => value - 100 >= 2 && value - 100 <= 2.5));
            Assert.IsTrue(negative.RawValues.All(value => 100 - value >= 7 && 100 - value <= 8));
        }

        [TestMethod]
        public void GreaterThanRequirementSupportsNegativeInternalInterval()
        {
            var generator = new MeasurementValueGenerator(new GenerationConfiguration(), new Random(7));
            var result = generator.Generate(new MeasurementGenerationInput
            {
                StandardValue = 10,
                Mpe = 0.5,
                ErrorType = ErrorType.Absolute,
                RequirementOperator = TechnicalRequirementOperator.GreaterThanOrEqual,
                ForceNegativeDirection = true,
                ValueCount = 3,
                DecimalPlaces = 3
            });

            Assert.IsTrue(result.RawValues.All(value => value < 10));
            Assert.IsTrue(result.RawValues.All(value => Math.Abs(value - 10) >= 0.5));
            Assert.IsTrue(result.RawValues.All(value => Math.Abs(value - 10) <= 0.65));
        }

        [DataTestMethod]
        [DataRow(0d, 0.001d, 6)]
        [DataRow(100000d, 5000d, 0)]
        public void MeasurementGroupFluctuationScalesWithRequirement(double standardValue, double mpe, int decimalPlaces)
        {
            var generator = new MeasurementValueGenerator(new GenerationConfiguration(), new Random(53));
            var result = generator.Generate(new MeasurementGenerationInput
            {
                StandardValue = standardValue,
                Mpe = mpe,
                ErrorType = ErrorType.Absolute,
                ForcePositiveDirection = true,
                ValueCount = 4,
                DecimalPlaces = decimalPlaces
            });
            var errors = result.RawValues.Select(value => value - standardValue).ToList();

            Assert.AreEqual(errors.Count, errors.Distinct().Count());
            Assert.IsTrue(errors.All(error => error > 0 && error <= mpe * 0.8 + 1e-9));
            Assert.IsTrue(errors.Max() - errors.Min() <= mpe * 0.06 + Math.Pow(10, -decimalPlaces));
        }

        [TestMethod]
        public void GenerationAlwaysUsesTemplateResolution()
        {
            const double maximumSpread = 0.3;
            var configuration = new GenerationConfiguration
            {
                UseDecimalPlacesForResolution = false,
                MeasurementGroupMaximumFluctuationCoefficient = maximumSpread
            };
            var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                new MeasurementValueGenerator(configuration, new Random(19)).Generate(new MeasurementGenerationInput
                {
                    StandardValue = 50,
                    Mpe = 1,
                    ErrorType = ErrorType.Absolute,
                    DistributionMode = DistributionMode.Uniform,
                    ForcePositiveDirection = true,
                    ValueCount = 6,
                    DecimalPlaces = 0,
                    MeasurementLowerBound = 0,
                    MeasurementUpperBound = 100
                }));

            StringAssert.Contains(exception.Message, "visibly different repeated measurements");
        }

        [TestMethod]
        public void GenerationRejectsGroupWhenResolutionCannotRepresentVisibleVariation()
        {
            var generator = new MeasurementValueGenerator(new GenerationConfiguration(), new Random(13));

            var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                generator.Generate(new MeasurementGenerationInput
                {
                    StandardValue = 1,
                    Mpe = 0.01,
                    ErrorType = ErrorType.Absolute,
                    ForcePositiveDirection = true,
                    ValueCount = 2,
                    DecimalPlaces = 2
                }));

            StringAssert.Contains(exception.Message, "visibly different repeated measurements");
        }

        [DataTestMethod]
        [DataRow(10d)]
        [DataRow(40d)]
        [DataRow(60d)]
        public void IntegerMeasurementsKeepVisibleVariationForEachStandardValue(double standardValue)
        {
            var result = new MeasurementValueGenerator(new GenerationConfiguration(), new Random(13)).Generate(
                new MeasurementGenerationInput
                {
                    StandardValue = standardValue,
                    Mpe = 5,
                    ErrorType = ErrorType.Absolute,
                    ForcePositiveDirection = true,
                    ValueCount = 3,
                    DecimalPlaces = 0
                });
            var errors = result.RawValues.Select(value => value - standardValue).ToList();

            Assert.IsTrue(result.DisplayValues.Distinct().Count() > 1);
            Assert.IsTrue(errors.Max() - errors.Min() <= 1 + 1e-12);
            Assert.IsTrue(errors.All(error => error > 0 && error <= 5));
        }

        [TestMethod]
        public void GenerationUsesEachMeasurementCellsDecimalPlaces()
        {
            var configuration = new GenerationConfiguration
            {
                MeasurementGroupMinimumFluctuationCoefficient = 0.2,
                MeasurementGroupMaximumFluctuationCoefficient = 0.3
            };
            var result = new MeasurementValueGenerator(configuration, new Random(19)).Generate(new MeasurementGenerationInput
            {
                StandardValue = 10,
                Mpe = 1,
                ErrorType = ErrorType.Absolute,
                ForcePositiveDirection = true,
                ValueCount = 3,
                DecimalPlaces = 3,
                DecimalPlacesByValue = new List<int> { 1, 2, 3 }
            });

            Assert.AreEqual(result.RawValues[0], Math.Round(result.RawValues[0], 1), 1e-12);
            Assert.AreEqual(result.RawValues[1], Math.Round(result.RawValues[1], 2), 1e-12);
            Assert.AreEqual(result.RawValues[2], Math.Round(result.RawValues[2], 3), 1e-12);
            CollectionAssert.AreEqual(new[] { 1, 2, 3 }, result.DisplayValues.Select(value => value.Split('.').Last().Length).ToArray());
        }

        [TestMethod]
        public void ParameterParserTreatsFullWidthPercentAsRatio()
        {
            var parser = new ParameterValueParser();

            Assert.AreEqual(0.005, parser.ParseMpe("±0.5％", percentageAsRatio: true), 0.0000001);
        }

        [TestMethod]
        public void DraftBuilderRecognizesSpecificationFieldVocabulary()
        {
            var snapshot = new WorkbookSnapshot
            {
                WorkbookName = "book.xlsx",
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 1, Column = 1, Text = "示值误差" },
                            new CellMeta { Row = 2, Column = 1, Text = "标准输入" },
                            new CellMeta { Row = 2, Column = 2, Text = "采样值" },
                            new CellMeta { Row = 2, Column = 3, Text = "误差限" },
                            new CellMeta { Row = 2, Column = 4, Text = "Full Scale" },
                            new CellMeta { Row = 2, Column = 5, Text = "扩展不确定度" },
                            new CellMeta { Row = 2, Column = 6, Text = "合格判定" },
                            new CellMeta { Row = 3, Column = 1, Text = "10" },
                            new CellMeta { Row = 3, Column = 2, Text = string.Empty },
                            new CellMeta { Row = 3, Column = 3, Text = "±0.5" },
                            new CellMeta { Row = 3, Column = 4, Text = "100" },
                            new CellMeta { Row = 3, Column = 5, Formula = "=C3/2" },
                            new CellMeta { Row = 3, Column = 6, Formula = "=IF(ABS(C3)<=0.5,\"合格\",\"不合格\")" }
                        }
                    }
                }
            };
            var recognition = new RecognitionResult
            {
                Snapshot = snapshot,
                RecognizedFields = new List<RecognizedField>
                {
                    new RecognizedField
                    {
                        Alias = "示值误差",
                        Score = 96,
                        Range = new CellRange
                        {
                            SheetName = "Sheet1",
                            StartRow = 1,
                            EndRow = 3,
                            StartColumn = 1,
                            EndColumn = 6
                        }
                    }
                }
            };

            var mappings = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter()).BuildMappings(recognition);

            Assert.AreEqual(1, mappings.Count);
            Assert.IsNotNull(mappings[0].StandardValueRange);
            Assert.IsNotNull(mappings[0].MeasurementValueRange);
            Assert.IsNotNull(mappings[0].TechnicalRequirementRange);
            Assert.IsNotNull(mappings[0].RangeValueRange);
            Assert.IsNotNull(mappings[0].UncertaintyRange);
            Assert.IsNotNull(mappings[0].ResultRange);

            var rule = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter())
                .BuildDraftRules(recognition, mappings)
                .Single();
            Assert.IsNotNull(rule.UncertaintySource);
            Assert.IsNotNull(rule.ResultSource);
        }

        [DataTestMethod]
        [DataRow("标准器测量值")]
        [DataRow("标准值")]
        public void DraftBuilderCapturesSetpointStandardAndMeasurementPositionsWithoutRenamingLegacyRoles(string standardHeader)
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 1, Column = 1, Text = "示值误差" },
                            new CellMeta { Row = 2, Column = 1, Text = "设定值" },
                            new CellMeta { Row = 2, Column = 2, Text = standardHeader },
                            new CellMeta { Row = 2, Column = 3, Text = "仪器测量值" },
                            new CellMeta { Row = 2, Column = 4, Text = "示值误差" },
                            new CellMeta { Row = 2, Column = 5, Text = "技术要求" },
                            new CellMeta { Row = 3, Column = 1, Text = "100.0", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 2, Text = "99.8", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 3, Text = "100.4", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 4, Text = "0.6", Formula = "=C3-B3", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 5, Text = "±1.0" }
                        }
                    }
                }
            };
            var recognition = RecognitionFor(snapshot, 5);
            var builder = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter());

            var mapping = builder.BuildMappings(recognition).Single();

            Assert.AreEqual(1, mapping.SetpointValueRange.StartColumn);
            Assert.AreEqual(2, mapping.StandardValueRange.StartColumn);
            Assert.AreEqual(3, mapping.MeasurementValueRange.StartColumn);

            var rule = builder.BuildDraftRules(recognition, new[] { mapping }).Single();
            new MeasurementRuleStructureAnalyzer().Apply(snapshot, new[] { rule });
            new RowMappingBuilder().Apply(snapshot, new[] { rule });

            Assert.AreEqual("设定值", rule.SetpointSource.Name);
            Assert.AreEqual("标准值", rule.StandardValueSource.Name);
            Assert.AreEqual(2, rule.StandardValueSource.Range.StartColumn);
            Assert.AreEqual(3, rule.TargetRange.StartColumn);
            Assert.IsTrue(rule.ErrorFormula.ReferencesStandardValue);
            Assert.AreEqual(1, rule.RowMappings.Single().SetpointValueRange.StartColumn);
            Assert.AreEqual(2, rule.RowMappings.Single().StandardValueRange.StartColumn);
            Assert.IsTrue(rule.TemplateDefinition.Regions.Any(region => region.Role == TemplateRegionRole.SetpointValue));
        }

        [TestMethod]
        public void DraftBuilderKeepsLegacySetpointHeaderAsStandardValueWhenThirdFieldIsAbsent()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 1, Column = 1, Text = "示值误差" },
                            new CellMeta { Row = 2, Column = 1, Text = "设定值" },
                            new CellMeta { Row = 2, Column = 2, Text = "测量值" },
                            new CellMeta { Row = 2, Column = 3, Text = "示值误差" },
                            new CellMeta { Row = 2, Column = 4, Text = "技术要求" },
                            new CellMeta { Row = 3, Column = 1, Text = "100.0", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 2, Text = "100.4", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 3, Text = "0.4", Formula = "=B3-A3", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 4, Text = "±1.0" }
                        }
                    }
                }
            };

            var mapping = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter())
                .BuildMappings(RecognitionFor(snapshot, 4))
                .Single();

            Assert.IsNull(mapping.SetpointValueRange);
            Assert.AreEqual(1, mapping.StandardValueRange.StartColumn);
            Assert.AreEqual(2, mapping.MeasurementValueRange.StartColumn);
        }

        private static RecognitionResult RecognitionFor(WorkbookSnapshot snapshot, int endColumn)
        {
            return new RecognitionResult
            {
                Snapshot = snapshot,
                RecognizedFields = new List<RecognizedField>
                {
                    new RecognizedField
                    {
                        Alias = "示值误差",
                        Score = 96,
                        Range = new CellRange
                        {
                            SheetName = "Sheet1",
                            StartRow = 1,
                            EndRow = 3,
                            StartColumn = 1,
                            EndColumn = endColumn
                        }
                    }
                }
            };
        }

        [TestMethod]
        public void DraftBuilderSkipsFormulaBasedErrorUnitRow()
        {
            var snapshot = new WorkbookSnapshot
            {
                WorkbookName = "book.xlsx",
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 1, Column = 1, Text = "\u793A\u503C\u8BEF\u5DEE" },
                            new CellMeta { Row = 2, Column = 1, Text = "\u6807\u51C6\u503C" },
                            new CellMeta { Row = 2, Column = 3, Text = "\u6D4B\u91CF\u503C" },
                            new CellMeta { Row = 2, Column = 6, Text = "\u793A\u503C\u8BEF\u5DEE\uFF08%FS\uFF09" },
                            new CellMeta { Row = 2, Column = 7, Text = "\u6280\u672F\u8981\u6C42\uFF08%FS\uFF09" },
                            new CellMeta { Row = 3, Column = 3, Text = "1" },
                            new CellMeta { Row = 3, Column = 4, Text = "2" },
                            new CellMeta { Row = 3, Column = 5, Text = "3" },
                            new CellMeta { Row = 3, Column = 6, Text = "%FS", Formula = "=IF(A4>0,\"%FS\",\"%\")" },
                            new CellMeta { Row = 4, Column = 1, Text = "10" },
                            new CellMeta { Row = 4, Column = 3, NumberFormat = "0.0" },
                            new CellMeta { Row = 4, Column = 4, NumberFormat = "0.0" },
                            new CellMeta { Row = 4, Column = 5, NumberFormat = "0.0" },
                            new CellMeta { Row = 4, Column = 6, Text = "0.0", Formula = "=AVERAGE(C4:E4)-A4" },
                            new CellMeta { Row = 4, Column = 7, Text = "\u00B13" }
                        }
                    }
                }
            };
            var recognition = new RecognitionResult
            {
                Snapshot = snapshot,
                RecognizedFields = new List<RecognizedField>
                {
                    new RecognizedField
                    {
                        Alias = "\u793A\u503C\u8BEF\u5DEE",
                        Score = 96,
                        Range = new CellRange
                        {
                            SheetName = "Sheet1",
                            StartRow = 1,
                            EndRow = 4,
                            StartColumn = 1,
                            EndColumn = 7
                        }
                    }
                }
            };

            var mapping = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter())
                .BuildMappings(recognition)
                .Single();

            Assert.AreEqual(4, mapping.MeasurementValueRange.StartRow);
            Assert.AreEqual(4, mapping.ErrorValueRange.StartRow);
        }

        [TestMethod]
        public void DraftBuilderUsesResponseTimeAverageAsErrorRange()
        {
            var sheet = new SheetSnapshot
            {
                Name = "Sheet1",
                Cells = new List<CellMeta>
                {
                    new CellMeta { Row = 1, Column = 1, Text = "\u54CD\u5E94\u65F6\u95F4" },
                    new CellMeta
                    {
                        Row = 2,
                        Column = 1,
                        Text = "\u6D4B\u91CF\u503C",
                        MergeRange = new CellRange { SheetName = "Sheet1", StartRow = 2, EndRow = 2, StartColumn = 1, EndColumn = 3 }
                    },
                    new CellMeta { Row = 2, Column = 4, Text = "\u5E73\u5747\u503C" },
                    new CellMeta { Row = 2, Column = 5, Text = "\u6280\u672F\u8981\u6C42" },
                    new CellMeta { Row = 3, Column = 1, NumberFormat = "0.0" },
                    new CellMeta { Row = 3, Column = 2, NumberFormat = "0.0" },
                    new CellMeta { Row = 3, Column = 3, NumberFormat = "0.0" },
                    new CellMeta { Row = 3, Column = 4, Text = "0.0", Formula = "=AVERAGE(A3:C3)", NumberFormat = "0.0" },
                    new CellMeta { Row = 3, Column = 5, Text = "<30" }
                }
            };
            var recognition = new RecognitionResult
            {
                Snapshot = new WorkbookSnapshot { Sheets = new List<SheetSnapshot> { sheet } },
                RecognizedFields = new List<RecognizedField>
                {
                    new RecognizedField
                    {
                        Alias = "\u54CD\u5E94\u65F6\u95F4",
                        Score = 96,
                        Range = new CellRange
                        {
                            SheetName = "Sheet1",
                            StartRow = 1,
                            EndRow = 3,
                            StartColumn = 1,
                            EndColumn = 5
                        }
                    }
                }
            };
            var builder = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter());

            var mapping = builder.BuildMappings(recognition).Single();
            var rule = builder.BuildDraftRules(recognition, new[] { mapping }).Single();

            Assert.IsNotNull(mapping.AverageValueRange);
            Assert.IsNotNull(mapping.ErrorValueRange);
            Assert.AreEqual(4, mapping.AverageValueRange.StartColumn);
            Assert.AreEqual(mapping.AverageValueRange.StartColumn, mapping.ErrorValueRange.StartColumn);
            Assert.AreEqual(mapping.AverageValueRange.EndColumn, mapping.ErrorValueRange.EndColumn);
            Assert.AreEqual(rule.AverageSource.Range.StartColumn, rule.ErrorSource.Range.StartColumn);
        }

        [TestMethod]
        public void FieldMatcherDoesNotTreatTableHeaderAsNestedSection()
        {
            var sheet = new SheetSnapshot
            {
                Name = "Sheet1",
                Cells = new List<CellMeta>
                {
                    new CellMeta { Row = 5, Column = 1, Text = "\u56db\u3001\u62a5\u8b66\u529f\u80fd\u53ca\u62a5\u8b66\u52a8\u4f5c\u503c" },
                    new CellMeta { Row = 6, Column = 1, Text = "\u62a5\u8b66\u529f\u80fd" },
                    new CellMeta { Row = 6, Column = 3, Text = "\u5b9e\u6d4b\u62a5\u8b66\u503c" },
                    new CellMeta { Row = 6, Column = 5, Text = "\u62a5\u8b66\u52a8\u4f5c\u503c" },
                    new CellMeta { Row = 7, Column = 1, Text = "\u58f0\u5149\u62a5\u8b66\u6b63\u5e38" },
                    new CellMeta { Row = 7, Column = 3, Text = "25" }
                }
            };

            var fields = new FieldMatcher().MatchMeasurementFields(sheet);

            Assert.AreEqual(1, fields.Count);
            Assert.AreEqual(5, fields[0].Range.StartRow);
        }

        [TestMethod]
        public void MergedRangeIsReturnedAsOneLogicalCell()
        {
            var merge = new CellRange { SheetName = "Sheet1", StartRow = 2, StartColumn = 1, EndRow = 3, EndColumn = 2 };
            var sheet = new SheetSnapshot
            {
                Name = "Sheet1",
                Cells = new List<CellMeta>
                {
                    new CellMeta { Row = 2, Column = 1, Text = "技术要求", IsMerged = true, MergeRange = merge },
                    new CellMeta { Row = 2, Column = 2, IsMerged = true, MergeRange = merge },
                    new CellMeta { Row = 3, Column = 1, IsMerged = true, MergeRange = merge },
                    new CellMeta { Row = 3, Column = 2, IsMerged = true, MergeRange = merge }
                }
            };

            var cells = MergedCellLogicalRangeResolver.GetTextCells(sheet, merge);

            Assert.AreEqual(1, cells.Count);
            Assert.AreEqual("技术要求", cells[0].Anchor.Text);
        }

        [TestMethod]
        public void GenerationDoesNotRequireHistoricalErrorFormula()
        {
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = Range("C5:C5"),
                ErrorSource = new ParameterSource { Range = Range("D5:D5") },
                FixedStandardValue = 10,
                FixedMpe = 0.5,
                WritableCells = new List<CellAddress> { new CellAddress { Row = 5, Column = 3 } }
            };

            GenerationRuleValidator.ValidateRule(rule, 1);

            Assert.AreEqual(1, rule.GroupSize);
        }

        [TestMethod]
        public void AlarmTemplateCanBeSavedWithOnlyMeasurementRange()
        {
            var repository = CreateRepository();
            var rule = new MeasurementRule
            {
                FieldName = "报警校准项",
                TargetRange = Range("C5:F5")
            };

            repository.ValidateTemplateForSave(new[] { rule });
        }

        [TestMethod]
        public void AlarmOverrideAppliesValueWithoutSelectingCalibrationItem()
        {
            var rules = new[]
            {
                new MeasurementRule { FieldName = "报警功能", FixedStandardValue = null },
                new MeasurementRule { FieldName = "示值误差", FixedStandardValue = 10 }
            };

            var overridden = ExcelCalibrationAddin.Host.Vsto.VstoAddinFacade.ApplyGenerationOverride(
                rules,
                new MeasurementGenerationOverride { AlarmValue = 2.5 });

            Assert.AreEqual(2.5, overridden[0].FixedStandardValue.Value, 1e-12);
            Assert.AreEqual(10, overridden[1].FixedStandardValue.Value, 1e-12);
        }

        [TestMethod]
        public void AlarmGenerationWritesEnteredValueToEveryTargetCell()
        {
            var writer = new RecordingWorkbookWriter();
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(12)),
                new GenerationConfiguration(),
                writer,
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "报警校准项",
                TargetRange = Range("C5:F5"),
                FixedStandardValue = 2,
                WritableCells = new List<CellAddress>
                {
                    new CellAddress { Row = 5, Column = 3 },
                    new CellAddress { Row = 5, Column = 4 },
                    new CellAddress { Row = 5, Column = 5 },
                    new CellAddress { Row = 5, Column = 6 }
                }
            };

            useCase.WritePreResolved(new[] { rule });

            Assert.AreEqual(1, writer.WriteCount);
            CollectionAssert.AreEqual(new[] { "2", "2", "2", "2" }, writer.LastValues.ToArray());
        }

        [TestMethod]
        public void MaximumErrorTemplateCanBeSavedWithoutPerPointErrorMappings()
        {
            var repository = CreateRepository();
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = RangeSpan(5, 6, 3),
                StandardValueSource = new ParameterSource { Range = RangeSpan(5, 6, 1) },
                ErrorSource = new ParameterSource { Range = RangeAt(8, 4) },
                MpeSource = new ParameterSource { Range = RangeSpan(5, 6, 5) },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=MAX(D5:D6)"
                }
            };

            repository.ValidateTemplateForSave(new[] { rule });
        }

        [TestMethod]
        public void MaximumErrorFormulaMakesEveryMeasurementRowMappingComplete()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "10", RawValueText = "10" },
                            new CellMeta { Row = 6, Column = 1, Text = "20", RawValueText = "20" },
                            new CellMeta { Row = 5, Column = 3 },
                            new CellMeta { Row = 6, Column = 3 },
                            new CellMeta { Row = 5, Column = 5, Text = "±1", RawValueText = "±1" },
                            new CellMeta { Row = 6, Column = 5, Text = "±1", RawValueText = "±1" },
                            new CellMeta { Row = 8, Column = 4, Formula = "=MAX(D5:D6)" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = RangeSpan(5, 6, 3),
                StandardValueSource = new ParameterSource { Range = RangeSpan(5, 6, 1) },
                ErrorSource = new ParameterSource { Range = RangeAt(8, 4) },
                MpeSource = new ParameterSource { Range = RangeSpan(5, 6, 5) },
                WritableCells = new List<CellAddress>
                {
                    new CellAddress { Row = 5, Column = 3 },
                    new CellAddress { Row = 6, Column = 3 }
                },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=MAX(D5:D6)"
                }
            };

            var mappings = new RowMappingBuilder().Build(snapshot, rule);

            Assert.AreEqual(2, mappings.Count);
            Assert.IsTrue(mappings.All(mapping => mapping.IsComplete), string.Join(";", mappings.Select(mapping => mapping.StatusMessage)));
            Assert.IsTrue(mappings.All(mapping => mapping.ErrorRange == null));
        }

        [TestMethod]
        public void MaximumErrorOnlyTemplateGeneratesEveryStandardPoint()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "10", RawValueText = "10" },
                            new CellMeta { Row = 6, Column = 1, Text = "20", RawValueText = "20" },
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0.00" },
                            new CellMeta
                            {
                                Row = 8,
                                Column = 4,
                                Text = "0",
                                RawValueText = "0",
                                Formula = "=MAX(ABS(C5-A5),ABS(C6-A6))",
                                NumberFormat = "0.00"
                            }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = RangeSpan(5, 6, 3),
                StandardValueSource = new ParameterSource { Range = RangeSpan(5, 6, 1) },
                ErrorSource = new ParameterSource { Range = RangeAt(8, 4) },
                FixedMpe = 1,
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=MAX(ABS(C5-A5),ABS(C6-A6))"
                }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(13)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.AreEqual(2, preview.RawValues.Count);
            Assert.IsTrue(Math.Abs(preview.RawValues[0] - 10) <= 1 + 1e-12);
            Assert.IsTrue(Math.Abs(preview.RawValues[1] - 20) <= 1 + 1e-12);
        }

        [TestMethod]
        public void StabilityMaxMinusMinFormulaIsNotClassifiedAsMaximumError()
        {
            var stabilityRule = new MeasurementRule
            {
                FieldName = "稳定性",
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=MAX(C5:C6)-MIN(C5:C6)",
                    ReferencesMeasurement = true
                }
            };

            Assert.IsFalse(ErrorFormulaClassifier.IsMaximumError(stabilityRule));

            var maximumErrorRule = new MeasurementRule
            {
                FieldName = "示值误差",
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=MAX(D5:D6)-MIN(D5:D6)"
                }
            };

            Assert.IsTrue(ErrorFormulaClassifier.IsMaximumError(maximumErrorRule));
        }

        [TestMethod]
        public void GenerationSkipsEmptyStandardValueRows()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "10" },
                            new CellMeta { Row = 6, Column = 1, Text = string.Empty },
                            new CellMeta { Row = 5, Column = 3, Text = string.Empty },
                            new CellMeta { Row = 6, Column = 3, Text = string.Empty }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "row-mapped",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 3, EndColumn = 3 },
                StandardValueSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 1, EndColumn = 1 }
                },
                FixedStandardValue = 10,
                FixedMpe = 0.5,
                FormatRule = new FormatRule { DecimalPlaces = 2 }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(7)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.Preview(new[] { rule }).Single();

            Assert.AreEqual(1, preview.WritableCells.Count);
            Assert.AreEqual(5, preview.WritableCells[0].Row);
            Assert.AreEqual(1, preview.DisplayValues.Count);
        }

        [TestMethod]
        public void MultipleStandardsUseDistinctErrorsWithinConfiguredTrend()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "10" },
                            new CellMeta { Row = 6, Column = 1, Text = "20" },
                            new CellMeta { Row = 7, Column = 1, Text = "30" },
                            new CellMeta { Row = 5, Column = 3, Text = string.Empty, NumberFormat = "0.00" },
                            new CellMeta { Row = 6, Column = 3, Text = string.Empty, NumberFormat = "0.00" },
                            new CellMeta { Row = 7, Column = 3, Text = string.Empty, NumberFormat = "0.00" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "trend",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 3, EndColumn = 3 },
                StandardValueSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 1, EndColumn = 1 }
                },
                FixedStandardValue = 10,
                FixedMpe = 5,
                FormatRule = new FormatRule { DecimalPlaces = 2 }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(61)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            var standards = new[] { 10d, 20d, 30d };
            var errors = preview.RawValues.Select((value, index) => Math.Round(value - standards[index], 2)).ToList();

            Assert.AreEqual(3, errors.Distinct().Count());
            Assert.IsTrue(errors.All(error => Math.Abs(error) >= 1 && Math.Abs(error) <= 4));
            Assert.IsTrue(errors.Skip(1).All(error => Math.Abs(Math.Abs(error) - Math.Abs(errors[0])) <= 1.01));
        }

        [TestMethod]
        public void SeparateRulesForOneItemKeepErrorsDistinctAcrossStandards()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "10" },
                            new CellMeta { Row = 6, Column = 1, Text = "40" },
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0" }
                        }
                    }
                }
            };
            var first = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 3, EndColumn = 3 },
                FixedStandardValue = 10,
                FixedMpe = 5,
                PositiveDirectionOnly = true,
                FormatRule = new FormatRule { DecimalPlaces = 1 }
            };
            var second = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 6, EndRow = 6, StartColumn = 3, EndColumn = 3 },
                FixedStandardValue = 40,
                FixedMpe = 5,
                PositiveDirectionOnly = true,
                FormatRule = new FormatRule { DecimalPlaces = 0 }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(61)),
                new GenerationConfiguration { UseDecimalPlacesForResolution = true },
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var previews = useCase.PreviewPreResolved(new[] { first, second });
            var errors = previews.Select((preview, index) =>
                preview.RawValues.Single() - new[] { 10d, 40d }[index]).ToList();

            Assert.AreEqual(2, errors.Distinct().Count());
        }

        [TestMethod]
        public void StandardGenerationRejectsErrorRoundedToZero()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.0" },
                            new CellMeta { Row = 5, Column = 4, Formula = "=C5-A5", NumberFormat = "0" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "zero rounded error",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 3, EndColumn = 3 },
                ErrorSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 4, EndColumn = 4 }
                },
                FixedStandardValue = 10,
                FixedMpe = 0.4,
                PositiveDirectionOnly = true,
                FormatRule = new FormatRule { DecimalPlaces = 1 }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(62)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                useCase.PreviewPreResolved(new[] { rule }));

            StringAssert.Contains(exception.Message, "无法生成有效误差值");
        }

        [TestMethod]
        public void MultipleStandardsRejectDuplicateErrorWhenResolutionHasOnlyOneValidStep()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "10" },
                            new CellMeta { Row = 6, Column = 1, Text = "20" },
                            new CellMeta { Row = 7, Column = 1, Text = "30" },
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0" },
                            new CellMeta { Row = 7, Column = 3, NumberFormat = "0" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "resolution constrained trend",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 3, EndColumn = 3 },
                StandardValueSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 1, EndColumn = 1 }
                },
                FixedStandardValue = 10,
                FixedMpe = 2,
                PositiveDirectionOnly = true,
                FormatRule = new FormatRule { DecimalPlaces = 0 }
            };
            var configuration = new GenerationConfiguration
            {
                UseSameDeviationDirection = true,
                PositiveErrorMinimumCoefficient = 0.5,
                PositiveErrorMaximumCoefficient = 0.6,
                ResultGroupFluctuationCoefficient = 0.3
            };
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(29)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var exception = Assert.ThrowsException<InvalidOperationException>(
                () => useCase.PreviewPreResolved(new[] { rule }));

            StringAssert.Contains(exception.Message, "不同标准值必须产生不同误差");
        }

        [TestMethod]
        public void MultipleStandardsUseErrorAreaPrecisionForGlobalTrend()
        {
            var cells = new List<CellMeta>
            {
                new CellMeta { Row = 1, Column = 8, Text = "0" },
                new CellMeta { Row = 1, Column = 9, Text = "100" }
            };
            var standards = new[] { 10d, 40d, 60d };
            for (var index = 0; index < standards.Length; index++)
            {
                var row = index + 5;
                cells.Add(new CellMeta { Row = row, Column = 1, Text = standards[index].ToString(CultureInfo.InvariantCulture) });
                cells.Add(new CellMeta { Row = row, Column = 3, NumberFormat = "0.0" });
                cells.Add(new CellMeta { Row = row, Column = 4, NumberFormat = "0.0" });
                cells.Add(new CellMeta { Row = row, Column = 5, NumberFormat = "0.0" });
                cells.Add(new CellMeta { Row = row, Column = 6, Formula = $"=AVERAGE(C{row}:E{row})", NumberFormat = "0.0" });
                cells.Add(new CellMeta
                {
                    Row = row,
                    Column = 7,
                    Formula = $"=(F{row}-A{row})/($I$1-$H$1)*100",
                    NumberFormat = "0.0"
                });
            }

            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot { Name = "Sheet1", Cells = cells }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "referenced trend",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 3, EndColumn = 5 },
                StandardValueSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 1, EndColumn = 1 }
                },
                AverageSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 6, EndColumn = 6 }
                },
                ErrorSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 7, EndColumn = 7 }
                },
                RangeSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 1, EndRow = 1, StartColumn = 8, EndColumn = 9 }
                },
                FixedStandardValue = standards[0],
                FixedMpe = 0.05,
                FixedReferenceRange = 100,
                ErrorType = ErrorType.Referenced,
                FormatRule = new FormatRule { DecimalPlaces = 1 }
            };
            var configuration = new GenerationConfiguration
            {
                UseSameDeviationDirection = true,
                UseIndependentDeviationControl = true,
                PositiveErrorMinimumCoefficient = 0.2,
                PositiveErrorMaximumCoefficient = 0.8,
                NegativeErrorMinimumCoefficient = 0.2,
                NegativeErrorMaximumCoefficient = 0.8,
                MeasurementGroupMaximumFluctuationCoefficient = 0.5,
                ResultGroupFluctuationCoefficient = 0.2
            };
            for (var seed = 0; seed < 200; seed++)
            {
                var useCase = new GenerateMeasurementUseCase(
                    current => new MeasurementValueGenerator(current, new Random(seed)),
                    configuration,
                    new RecordingWorkbookWriter(),
                    new StaticSnapshotProvider(snapshot),
                    null);

                var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
                var trendErrors = new List<double>();
                for (var index = 0; index < standards.Length; index++)
                {
                    var writtenValues = preview.DisplayValues
                        .Skip(index * 3)
                        .Take(3)
                        .Select(value => double.Parse(value, CultureInfo.InvariantCulture))
                        .ToList();
                    Assert.IsTrue(writtenValues.Max() - writtenValues.Min() <= 3);
                    trendErrors.Add(Math.Round(writtenValues.Average() - standards[index], 1));
                }

                Assert.AreEqual(3, trendErrors.Distinct().Count(), $"seed={seed}");
                Assert.IsTrue(trendErrors.All(error => Math.Abs(error) >= 1 && Math.Abs(error) <= 4), $"seed={seed}");
                Assert.IsTrue(trendErrors.All(error => Math.Sign(error) == Math.Sign(trendErrors[0])), $"seed={seed}");
                Assert.IsTrue(trendErrors.Max() - trendErrors.Min() <= 1 + 1e-12, $"seed={seed}");
            }
        }

        [TestMethod]
        public void PreResolvedReferencedScaleSurvivesPointFormulaReanalysis()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 1, Column = 8, Text = "0" },
                            new CellMeta { Row = 1, Column = 9, Text = "100" },
                            new CellMeta { Row = 5, Column = 1, Text = "10" },
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0" },
                            new CellMeta { Row = 5, Column = 4, NumberFormat = "0" },
                            new CellMeta { Row = 5, Column = 5, NumberFormat = "0" },
                            new CellMeta { Row = 5, Column = 6, Formula = "=AVERAGE(C5:E5)", NumberFormat = "0.0" },
                            new CellMeta { Row = 5, Column = 7, Formula = "=(F5-A5)/A5*100", NumberFormat = "0.0" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "referenced point scale",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 3, EndColumn = 5 },
                StandardValueSource = new ParameterSource { Range = RangeAt(5, 1) },
                AverageSource = new ParameterSource { Range = RangeAt(5, 6) },
                ErrorSource = new ParameterSource { Range = RangeAt(5, 7) },
                RangeSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 1, EndRow = 1, StartColumn = 8, EndColumn = 9 }
                },
                FixedStandardValue = 10,
                FixedMpe = 0.05,
                FixedReferenceRange = 100,
                ErrorType = ErrorType.Referenced,
                FormatRule = new FormatRule { DecimalPlaces = 0 },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=(F5-A5)/A5*100",
                    ReferencesAverage = true,
                    AverageFormulaResolved = true,
                    AverageFormula = "=AVERAGE(C5:E5)",
                    TechnicalRequirementFormula = "=J5",
                    ResultFormula = "=K5",
                    Scale = ErrorFormulaScale.RelativeToReferenceRange,
                    FormulaMultipliesBy100 = true,
                    FormulaDividesByReferenceRange = true
                }
            };
            var configuration = new GenerationConfiguration
            {
                UseIndependentDeviationControl = true,
                PositiveErrorMinimumCoefficient = 0.2,
                PositiveErrorMaximumCoefficient = 0.8,
                NegativeErrorMinimumCoefficient = 0.2,
                NegativeErrorMaximumCoefficient = 0.8,
                UseDecimalPlacesForResolution = false
            };
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(17)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            var generatedError = preview.RawValues.Average() - 10;

            Assert.AreEqual(ErrorFormulaScale.RelativeToReferenceRange, preview.Rule.ErrorFormula.Scale);
            Assert.IsTrue(Math.Abs(generatedError) >= 1 && Math.Abs(generatedError) <= 4);
        }

        [TestMethod]
        public void PercentageFormulaNormalizesUnitlessCachedMpeBeforeGeneration()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 1, Column = 8, Text = "0" },
                            new CellMeta { Row = 1, Column = 9, Text = "30" },
                            new CellMeta { Row = 5, Column = 1, Text = "6.0" },
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.0" },
                            new CellMeta { Row = 5, Column = 4, NumberFormat = "0.0" },
                            new CellMeta { Row = 5, Column = 5, NumberFormat = "0.0" },
                            new CellMeta { Row = 5, Column = 6, Formula = "=AVERAGE(C5:E5)", NumberFormat = "0.0" },
                            new CellMeta { Row = 5, Column = 7, Formula = "=(F5-A5)/($I$1-$H$1)*100", NumberFormat = "0.0" },
                            new CellMeta { Row = 5, Column = 10, Text = "±3.0" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "oxygen indication error",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 3, EndColumn = 5 },
                StandardValueSource = new ParameterSource { Range = RangeAt(5, 1) },
                AverageSource = new ParameterSource { Range = RangeAt(5, 6) },
                ErrorSource = new ParameterSource { Range = RangeAt(5, 7) },
                MpeSource = new ParameterSource
                {
                    Range = RangeAt(5, 10),
                    ValuePattern = "mpe:absolute:scale=1"
                },
                RangeSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 1, EndRow = 1, StartColumn = 8, EndColumn = 9 }
                },
                FixedStandardValue = 6,
                FixedMpe = 3,
                FixedReferenceRange = 30,
                ErrorType = ErrorType.Absolute,
                FormatRule = new FormatRule { DecimalPlaces = 1 },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=(F5-A5)/($I$1-$H$1)*100",
                    ReferencesAverage = true,
                    AverageFormulaResolved = true,
                    AverageFormula = "=AVERAGE(C5:E5)",
                    TechnicalRequirementFormula = "=J5",
                    ResultFormula = "=K5",
                    Scale = ErrorFormulaScale.RelativeToReferenceRange,
                    FormulaMultipliesBy100 = true,
                    FormulaDividesByReferenceRange = true
                }
            };
            var configuration = new GenerationConfiguration
            {
                UseIndependentDeviationControl = true,
                PositiveErrorMinimumCoefficient = 0.2,
                PositiveErrorMaximumCoefficient = 0.8,
                NegativeErrorMinimumCoefficient = 0.2,
                NegativeErrorMaximumCoefficient = 0.8,
                UseDecimalPlacesForResolution = false
            };
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(29)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            var generatedError = preview.RawValues.Average() - 6;

            Assert.AreEqual(0.03, preview.Rule.FixedMpe.Value, 1e-12);
            Assert.AreEqual(ErrorType.Referenced, preview.Rule.ErrorType);
            StringAssert.StartsWith(preview.Rule.MpeSource.ValuePattern, "mpe:referenced:scale=0.01");
            Assert.IsTrue(Math.Abs(generatedError) >= 0.1 && Math.Abs(generatedError) <= 0.8);
        }

        [TestMethod]
        public void RelativeTrendUsesRoundedAverageFormulaBeforeDisplayDeduplication()
        {
            var cells = new List<CellMeta>();
            var standards = new[] { 20d, 50d, 80d };
            for (var index = 0; index < standards.Length; index++)
            {
                var row = index + 5;
                cells.Add(new CellMeta { Row = row, Column = 1, Text = standards[index].ToString(CultureInfo.InvariantCulture) });
                cells.Add(new CellMeta { Row = row, Column = 3, NumberFormat = "0.0" });
                cells.Add(new CellMeta { Row = row, Column = 4, NumberFormat = "0.0" });
                cells.Add(new CellMeta { Row = row, Column = 5, NumberFormat = "0.0" });
                cells.Add(new CellMeta
                {
                    Row = row,
                    Column = 6,
                    Formula = $"=ROUND(AVERAGE(C{row}:E{row}),1)",
                    NumberFormat = "0.0"
                });
                cells.Add(new CellMeta
                {
                    Row = row,
                    Column = 7,
                    Formula = $"=(F{row}-A{row})/A{row}*100",
                    NumberFormat = "0.0"
                });
            }

            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot { Name = "Sheet1", Cells = cells }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "relative rounded-average trend",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 3, EndColumn = 5 },
                StandardValueSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 1, EndColumn = 1 }
                },
                AverageSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 6, EndColumn = 6 }
                },
                ErrorSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 7, EndColumn = 7 }
                },
                FixedStandardValue = standards[0],
                FixedMpe = 0.1,
                ErrorType = ErrorType.Relative,
                FormatRule = new FormatRule { DecimalPlaces = 0 }
            };
            var configuration = new GenerationConfiguration
            {
                UseSameDeviationDirection = true,
                UseDecimalPlacesForResolution = false,
                ResultGroupFluctuationCoefficient = 0.3
            };
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(15)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            var displayedErrors = new List<double>();
            for (var index = 0; index < standards.Length; index++)
            {
                var average = preview.RawValues.Skip(index * 3).Take(3).Average();
                var roundedAverage = Math.Round(average, 1, MidpointRounding.AwayFromZero);
                displayedErrors.Add(Math.Round(
                    (roundedAverage - standards[index]) / standards[index] * 100,
                    1,
                    MidpointRounding.AwayFromZero));
            }

            Assert.AreEqual(3, displayedErrors.Distinct().Count());
        }

        [TestMethod]
        public void SameStandardAcrossRulesReusesMeasurementSequence()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.000" },
                            new CellMeta { Row = 5, Column = 4, NumberFormat = "0.000" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0.000" },
                            new CellMeta { Row = 6, Column = 4, NumberFormat = "0.000" }
                        }
                    }
                }
            };
            var first = SharedStandardRule("first", 5);
            var second = SharedStandardRule("second", 6);
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(71)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var previews = useCase.PreviewPreResolved(new[] { first, second });

            CollectionAssert.AreEqual(previews[0].RawValues.ToArray(), previews[1].RawValues.ToArray());
            CollectionAssert.AreEqual(previews[0].DisplayValues.ToArray(), previews[1].DisplayValues.ToArray());
        }

        [TestMethod]
        public void SameStandardAcrossRulesUsesEachRulesMpeContract()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 5, Column = 4, NumberFormat = "0.00" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 6, Column = 4, NumberFormat = "0.00" }
                        }
                    }
                }
            };
            var largeMpeRule = SharedStandardRule("large MPE", 5);
            largeMpeRule.FixedMpe = 10;
            var smallMpeRule = SharedStandardRule("small MPE", 6);
            smallMpeRule.FixedMpe = 0.1;
            var configuration = new GenerationConfiguration
            {
                PositiveErrorMinimumCoefficient = 0.7,
                PositiveErrorMaximumCoefficient = 0.9,
                NegativeErrorMinimumCoefficient = 0.7,
                NegativeErrorMaximumCoefficient = 0.9
            };
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(74)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var previews = useCase.PreviewPreResolved(new[] { largeMpeRule, smallMpeRule });
            var smallMpeErrors = previews[1].RawValues.Select(value => Math.Abs(value - 10)).ToList();

            Assert.IsTrue(smallMpeErrors.All(error => error <= 0.1 + 1e-12));
            Assert.IsTrue(smallMpeErrors.All(error => Math.Abs(error - 0.08) <= 0.011));
        }

        [TestMethod]
        public void SmallMpeRejectsGroupWithoutRepresentableVariation()
        {
            var configuration = new GenerationConfiguration
            {
                PositiveErrorMinimumCoefficient = 0.2,
                PositiveErrorMaximumCoefficient = 0.8,
                NegativeErrorMinimumCoefficient = 0.2,
                NegativeErrorMaximumCoefficient = 0.8,
                UseDecimalPlacesForResolution = true
            };
            var generator = new MeasurementValueGenerator(configuration, new Random(75));

            var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                generator.Generate(new MeasurementGenerationInput
                {
                    StandardValue = 10,
                    Mpe = 0.01,
                    ErrorType = ErrorType.Absolute,
                    ValueCount = 3,
                    DecimalPlaces = 2
                }));

            StringAssert.Contains(exception.Message, "visibly different repeated measurements");
        }

        [TestMethod]
        public void ResolutionThatCannotRepresentNonZeroErrorFailsGeneration()
        {
            var generator = new MeasurementValueGenerator(new GenerationConfiguration(), new Random(76));

            var exception = Assert.ThrowsException<InvalidOperationException>(() => generator.Generate(
                new MeasurementGenerationInput
                {
                    StandardValue = 10,
                    Mpe = 0.1,
                    ErrorType = ErrorType.Absolute,
                    ForcePositiveDirection = true,
                    ValueCount = 1,
                    DecimalPlaces = 0
                }));

            StringAssert.Contains(exception.Message, "non-zero calibration error");
        }

        [TestMethod]
        public void CoarseResolutionStillProducesVisibleGroupVariation()
        {
            var configuration = new GenerationConfiguration
            {
                UseSameDeviationDirection = true,
                UseDecimalPlacesForResolution = true,
                MeasurementGroupMaximumFluctuationCoefficient = 0.06
            };

            for (var seed = 0; seed < 100; seed++)
            {
                var generator = new MeasurementValueGenerator(configuration, new Random(seed));
                var result = generator.Generate(new MeasurementGenerationInput
                {
                    StandardValue = 50,
                    Mpe = 0.1,
                    ErrorType = ErrorType.Relative,
                    ValueCount = 3,
                    DecimalPlaces = 1
                });
                var errors = result.RawValues.Select(value => value - 50).ToList();

                Assert.IsTrue(result.DisplayValues.Distinct().Count() > 1, $"seed={seed}");
                Assert.IsTrue(errors.Max() - errors.Min() <= 0.3 + 1e-12, $"seed={seed}");
                Assert.IsTrue(errors.All(error => Math.Abs(error) >= 1 - 1e-12 && Math.Abs(error) <= 4 + 1e-12), $"seed={seed}");
            }
        }

        [TestMethod]
        public void UnrepresentableMpeIntervalDoesNotReturnOutOfRangeFallback()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 5, Column = 4, Text = "0.00", NumberFormat = "0.00" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "unrepresentable MPE",
                TargetRange = RangeAt(5, 3),
                FixedStandardValue = 10,
                FixedMpe = 0.01,
                ErrorSource = new ParameterSource { Range = RangeAt(5, 4) },
                FormatRule = new FormatRule { DecimalPlaces = 2 }
            };
            var configuration = new GenerationConfiguration
            {
                PositiveErrorMinimumCoefficient = 0.2,
                PositiveErrorMaximumCoefficient = 0.8,
                NegativeErrorMinimumCoefficient = 0.2,
                NegativeErrorMaximumCoefficient = 0.8,
                UseDecimalPlacesForResolution = true
            };
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(76)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                useCase.PreviewPreResolved(new[] { rule }));

            StringAssert.Contains(exception.Message, "无法生成有效误差值");
        }

        [TestMethod]
        public void RepeatabilityUsesSharedStandardErrorAndKeepsVariation()
        {
            var cells = new List<CellMeta>();
            for (var column = 3; column <= 5; column++)
            {
                cells.Add(new CellMeta { Row = 5, Column = column, NumberFormat = "0" });
            }

            for (var column = 3; column <= 8; column++)
            {
                cells.Add(new CellMeta { Row = 6, Column = column, NumberFormat = "0" });
            }

            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot { Name = "Sheet1", Cells = cells }
                }
            };
            var indication = new MeasurementRule
            {
                FieldName = "indication error",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 3, EndColumn = 5 },
                FixedStandardValue = 40,
                FixedMpe = 5,
                PositiveDirectionOnly = true,
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    ReferencesAverage = true,
                    AverageFormulaResolved = true,
                    AverageFormula = "=AVERAGE(C5:E5)"
                },
                FormatRule = new FormatRule { DecimalPlaces = 0, UnitSuffix = "%LEL" }
            };
            var repeatability = new MeasurementRule
            {
                FieldName = "重复性",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 6, EndRow = 6, StartColumn = 3, EndColumn = 8 },
                FixedStandardValue = 40,
                FixedMpe = 2,
                FormatRule = new FormatRule { DecimalPlaces = 0, UnitSuffix = "%LEL" }
            };
            var configuration = new GenerationConfiguration
            {
                UseSameDeviationDirection = true,
                UseIndependentDeviationControl = true,
                PositiveErrorMinimumCoefficient = 0.6,
                PositiveErrorMaximumCoefficient = 0.8,
                NegativeErrorMinimumCoefficient = 0.6,
                NegativeErrorMaximumCoefficient = 0.8,
                AbsoluteErrorMinimumCoefficient = 0.2,
                AbsoluteErrorMaximumCoefficient = 0.8,
                UseDecimalPlacesForResolution = false,
                MeasurementGroupMaximumFluctuationCoefficient = 0.3,
                ResultGroupFluctuationCoefficient = 0.3
            };
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(83)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var previews = useCase.PreviewPreResolved(new[] { indication, repeatability });
            var indicationValues = previews[0].DisplayValues
                .Select(value => double.Parse(value, CultureInfo.InvariantCulture))
                .ToList();
            var repeatabilityValues = previews[1].DisplayValues
                .Select(value => double.Parse(value, CultureInfo.InvariantCulture))
                .ToList();
            var indicationAverage = indicationValues.Average();
            var repeatabilityAverage = repeatabilityValues.Average();
            var repeatabilityStandardDeviation = Math.Sqrt(repeatabilityValues
                .Sum(value => Math.Pow(value - repeatabilityAverage, 2)) / (repeatabilityValues.Count - 1));

            Assert.IsTrue(indicationAverage - 40 >= 3);
            Assert.IsTrue(Math.Abs(repeatabilityAverage - indicationAverage) <= 1);
            Assert.IsTrue(repeatabilityValues.Distinct().Count() > 1);
            Assert.IsTrue(repeatabilityStandardDeviation / Math.Abs(repeatabilityAverage) * 100 <= 2 + 1e-12);
        }

        [TestMethod]
        public void SameStandardAcrossRulesAllowsDifferentResolutions()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.000" },
                            new CellMeta { Row = 5, Column = 4, NumberFormat = "0.000" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 6, Column = 4, NumberFormat = "0.00" }
                        }
                    }
                }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(72)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var previews = useCase.PreviewPreResolved(new[]
            {
                SharedStandardRule("fine", 5),
                SharedStandardRule("coarse", 6)
            });

            Assert.AreEqual(2, previews.Count);
            Assert.IsTrue(previews.All(preview => preview.DisplayValues.Count == 2));
        }

        [TestMethod]
        public void SameStandardAcrossRulesRejectsConflictingForcedDirections()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.000" },
                            new CellMeta { Row = 5, Column = 4, NumberFormat = "0.000" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0.000" },
                            new CellMeta { Row = 6, Column = 4, NumberFormat = "0.000" }
                        }
                    }
                }
            };
            var first = SharedStandardRule("positive", 5);
            first.PositiveDirectionOnly = true;
            var second = SharedStandardRule("negative", 6);
            second.NegativeDirectionOnly = true;
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(73)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                useCase.PreviewPreResolved(new[] { first, second }));

            StringAssert.Contains(exception.Message, "方向");
        }

        [DataTestMethod]
        [DataRow("~")]
        [DataRow("～")]
        [DataRow("-")]
        [DataRow("–")]
        [DataRow("—")]
        [DataRow("至")]
        public void ParameterResolverUsesMappedRangeWhenFormulaRequiresReferencedError(string separator)
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "0.5" },
                            new CellMeta { Row = 6, Column = 1, Text = "0" },
                            new CellMeta { Row = 6, Column = 2, Text = separator },
                            new CellMeta { Row = 6, Column = 3, Text = "100" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "referenced error",
                TargetRange = RangeAt(5, 4),
                MpeSource = new ParameterSource { Range = RangeAt(5, 1) },
                RangeSource = new ParameterSource
                {
                    Range = new CellRange
                    {
                        SheetName = "Sheet1",
                        StartRow = 6,
                        EndRow = 6,
                        StartColumn = 1,
                        EndColumn = 3
                    }
                },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Scale = ErrorFormulaScale.RelativeToReferenceRange
                }
            };

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(100d, resolved.FixedReferenceRange);
        }

        [TestMethod]
        public void ParameterResolverTreatsHyphenAsRangeSeparator()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "0.5" },
                            new CellMeta { Row = 6, Column = 1, Text = "20-100" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "referenced error",
                TargetRange = RangeAt(5, 4),
                MpeSource = new ParameterSource { Range = RangeAt(5, 1) },
                RangeSource = new ParameterSource { Range = RangeAt(6, 1) },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Scale = ErrorFormulaScale.RelativeToReferenceRange
                }
            };

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(80d, resolved.FixedReferenceRange);
        }

        [TestMethod]
        public void RowMappingPreservesRangeValueRange()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "40" },
                            new CellMeta { Row = 5, Column = 2, Text = "0-100" },
                            new CellMeta { Row = 5, Column = 3 }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                TargetRange = RangeAt(5, 3),
                StandardValueSource = new ParameterSource { Range = RangeAt(5, 1) },
                RangeSource = new ParameterSource { Range = RangeAt(5, 2) },
                WritableCells = new List<CellAddress> { new CellAddress { Row = 5, Column = 3 } }
            };

            var mapping = new RowMappingBuilder().Build(snapshot, rule).Single();

            Assert.IsNotNull(mapping.RangeValueRange);
            Assert.AreEqual(5, mapping.RangeValueRange.StartRow);
            Assert.AreEqual(2, mapping.RangeValueRange.StartColumn);
        }

        [TestMethod]
        public void RowSpecificTechnicalRequirementUsesCurrentFormulaDisplayValue()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "10" },
                            new CellMeta { Row = 6, Column = 1, Text = "20" },
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.000" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0.000" },
                            new CellMeta { Row = 5, Column = 4, Text = "0", Formula = "=C5-A5" },
                            new CellMeta { Row = 6, Column = 4, Text = "0", Formula = "=C6-A6" },
                            new CellMeta { Row = 5, Column = 5, Text = "±0.5", DisplayText = "±0.5", Formula = "=IF(A5<15,\"±0.5\",\"±1.0\")" },
                            new CellMeta { Row = 6, Column = 5, Text = "±1.0", DisplayText = "±1.0", Formula = "=IF(A6<15,\"±0.5\",\"±1.0\")" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "row requirement",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 3, EndColumn = 3 },
                StandardValueSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 1, EndColumn = 1 } },
                ErrorSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 4, EndColumn = 4 } },
                MpeSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 5, EndColumn = 5 } },
                FixedStandardValue = 10,
                FixedMpe = 0.5,
                FormatRule = new FormatRule { DecimalPlaces = 3 }
            };
            var configuration = new GenerationConfiguration
            {
                PositiveErrorMinimumCoefficient = 0.8,
                PositiveErrorMaximumCoefficient = 0.8,
                NegativeErrorMinimumCoefficient = 0.8,
                NegativeErrorMaximumCoefficient = 0.8,
                ResultGroupFluctuationCoefficient = 0.5
            };
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(79)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                new MeasurementRuleParameterResolver());

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            var firstError = Math.Abs(preview.RawValues[0] - 10);
            var secondError = Math.Abs(preview.RawValues[1] - 20);

            Assert.AreEqual(0.4, firstError, 0.001);
            Assert.AreEqual(0.8, secondError, 0.001);
        }

        [TestMethod]
        public void RowSpecificDifferentMpeValuesPreserveConfiguredUsageRatio()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "10" },
                            new CellMeta { Row = 6, Column = 1, Text = "20" },
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.000" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0.000" },
                            new CellMeta { Row = 5, Column = 4, Text = "0", Formula = "=C5-A5", NumberFormat = "0.000" },
                            new CellMeta { Row = 6, Column = 4, Text = "0", Formula = "=C6-A6", NumberFormat = "0.000" },
                            new CellMeta { Row = 5, Column = 5, Text = "±0.1", DisplayText = "±0.1" },
                            new CellMeta { Row = 6, Column = 5, Text = "±10", DisplayText = "±10" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "different row MPE",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 3, EndColumn = 3 },
                StandardValueSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 1, EndColumn = 1 } },
                ErrorSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 4, EndColumn = 4 } },
                MpeSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 5, EndColumn = 5 } },
                FixedStandardValue = 10,
                FixedMpe = 0.1,
                FormatRule = new FormatRule { DecimalPlaces = 3 }
            };
            var configuration = new GenerationConfiguration
            {
                PositiveErrorMinimumCoefficient = 0.7,
                PositiveErrorMaximumCoefficient = 0.8,
                NegativeErrorMinimumCoefficient = 0.7,
                NegativeErrorMaximumCoefficient = 0.8,
                ResultGroupFluctuationCoefficient = 0.05
            };
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(80)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                new MeasurementRuleParameterResolver());

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            var usages = new[]
            {
                Math.Abs(preview.RawValues[0] - 10) / 0.1,
                Math.Abs(preview.RawValues[1] - 20) / 10
            };

            Assert.IsTrue(usages.All(usage => usage >= 0.7 - 1e-12 && usage <= 0.8 + 1e-12));
            Assert.IsTrue(Math.Abs(usages[0] - usages[1]) <= 0.05 + 0.011);
        }

        [TestMethod]
        public void RemoteSyncConflictCanUseRemote()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-use-remote");
            repository.SaveTemplate(
                "TemplateA",
                fingerprint,
                new[] { Rule("local", 10) },
                null,
                localSyncStatus: TemplateSyncStatus.PendingUpload);

            var applyResult = repository.UpsertRemoteTemplate(
                "remote-template-a",
                "TemplateA",
                2,
                DateTime.UtcNow,
                null,
                TemplateLifecycleStatus.Enabled,
                fingerprint,
                new[] { Rule("remote", 20) },
                null);

            Assert.AreEqual(TemplateRemoteApplyResult.Conflict, applyResult);
            Assert.IsTrue(repository.FindByExactFingerprint(fingerprint.ExactFingerprint).HasRemoteConflict);

            Assert.IsTrue(repository.ResolveTemplateConflict(
                fingerprint.ExactFingerprint,
                TemplateConflictResolutionAction.UseRemote));

            var resolved = repository.FindByExactFingerprint(fingerprint.ExactFingerprint);
            Assert.AreEqual(TemplateSyncStatus.Synced, resolved.LocalSyncStatus);
            Assert.AreEqual("remote-template-a", resolved.RemoteTemplateId);
            Assert.AreEqual(20, resolved.Rules[0].FixedStandardValue);
            Assert.IsFalse(resolved.HasRemoteConflict);
        }

        [TestMethod]
        public void UpdatingSameFingerprintTemplateUsesSelectedRemoteIdentity()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-shared-edit");
            repository.SaveTemplate("First", fingerprint, new[] { Rule("first", 10) }, null, createNew: true);
            repository.SaveTemplate("Second", fingerprint, new[] { Rule("second", 20) }, null, createNew: true);
            var first = repository.ListSavedTemplates().Single(item => item.TemplateName == "First");
            var second = repository.ListSavedTemplates().Single(item => item.TemplateName == "Second");

            repository.SaveTemplate(
                "First Updated",
                fingerprint,
                new[] { Rule("first updated", 30) },
                null,
                createNew: false,
                targetRemoteTemplateId: first.RemoteTemplateId);

            var templates = repository.ListSavedTemplates();
            Assert.AreEqual(2, templates.Count);
            Assert.AreEqual("First Updated", repository.FindByRemoteTemplateId(first.RemoteTemplateId).TemplateName);
            Assert.AreEqual(30, repository.FindByRemoteTemplateId(first.RemoteTemplateId).Rules[0].FixedStandardValue);
            Assert.AreEqual("Second", repository.FindByRemoteTemplateId(second.RemoteTemplateId).TemplateName);
        }

        [TestMethod]
        public void RemoteSyncConflictCanKeepLocal()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-keep-local");
            repository.SaveTemplate(
                "TemplateB",
                fingerprint,
                new[] { Rule("local", 10) },
                null,
                localSyncStatus: TemplateSyncStatus.PendingUpload);

            repository.UpsertRemoteTemplate(
                "remote-template-b",
                "TemplateB",
                2,
                DateTime.UtcNow,
                null,
                TemplateLifecycleStatus.Enabled,
                fingerprint,
                new[] { Rule("remote", 20) },
                null);

            Assert.IsTrue(repository.ResolveTemplateConflict(
                fingerprint.ExactFingerprint,
                TemplateConflictResolutionAction.KeepLocal));

            var resolved = repository.FindByExactFingerprint(fingerprint.ExactFingerprint);
            Assert.AreEqual(TemplateSyncStatus.PendingUpload, resolved.LocalSyncStatus);
            Assert.AreEqual(10, resolved.Rules[0].FixedStandardValue);
            Assert.IsFalse(resolved.HasRemoteConflict);
        }

        [TestMethod]
        public void RemoteSyncConflictCanSaveAs()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-save-as");
            repository.SaveTemplate(
                "TemplateC",
                fingerprint,
                new[] { Rule("local", 10) },
                null,
                localSyncStatus: TemplateSyncStatus.PendingUpload);

            repository.UpsertRemoteTemplate(
                "remote-template-c",
                "TemplateC",
                2,
                DateTime.UtcNow,
                null,
                TemplateLifecycleStatus.Enabled,
                fingerprint,
                new[] { Rule("remote", 20) },
                null);

            Assert.IsTrue(repository.ResolveTemplateConflict(
                fingerprint.ExactFingerprint,
                TemplateConflictResolutionAction.SaveAs,
                "TemplateC Local"));

            var templates = repository.ListSavedTemplates();
            Assert.IsTrue(templates.Any(item => item.TemplateName == "TemplateC" && item.LocalSyncStatus == TemplateSyncStatus.Synced));
            Assert.IsTrue(templates.Any(item => item.TemplateName == "TemplateC Local" && item.LocalSyncStatus == TemplateSyncStatus.PendingUpload));
        }

        [TestMethod]
        public void TemplateSaveWritesRemoteReturnedTemplateToLocalCache()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-remote-save");
            var remoteRule = Rule("remote accepted", 99);
            var response = JsonConvert.SerializeObject(new
            {
                ok = true,
                data = new
                {
                    template_id = "remote-save-1",
                    template_name = "Remote Accepted",
                    version = 7,
                    updated_at = DateTime.UtcNow.ToString("o"),
                    status = "enabled",
                    fingerprint,
                    rules = new[] { remoteRule }
                }
            });
            var client = new TemplateSyncClient(
                new HttpClient(new StaticJsonHandler(response)),
                "http://localhost:3002/api/templates");
            var service = new TemplateSaveService(client, repository);

            var result = service.Save(
                "Local Draft",
                fingerprint,
                new[] { Rule("local draft", 10) },
                null,
                createNew: false);

            var cached = repository.FindByExactFingerprint(fingerprint.ExactFingerprint);
            Assert.IsTrue(result.SavedToRemote);
            Assert.AreEqual(TemplateSyncStatus.Synced, cached.LocalSyncStatus);
            Assert.AreEqual("remote-save-1", cached.RemoteTemplateId);
            Assert.AreEqual("Remote Accepted", cached.TemplateName);
            Assert.AreEqual(7, cached.RemoteVersion);
            Assert.AreEqual(99, cached.Rules[0].FixedStandardValue);
        }

        [TestMethod]
        public void TemplateSavePersistsStructureButNotCurrentStandardOrRangeValues()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-static-template-definition");
            var rule = Rule("dynamic values", 10);
            rule.StandardValueSource = new ParameterSource { Name = "standard", Range = RangeAt(5, 1) };
            rule.RangeSource = new ParameterSource { Name = "range", Range = RangeAt(2, 9) };
            rule.FixedReferenceRange = 100;
            rule.TemplateDefinition = new TemplateFieldDefinition
            {
                ProjectName = "dynamic values",
                SectionRange = new CellRange { SheetName = "Sheet1", StartRow = 1, EndRow = 5, StartColumn = 1, EndColumn = 9 },
                Regions = new List<TemplateRegionDefinition>
                {
                    new TemplateRegionDefinition
                    {
                        Role = TemplateRegionRole.StandardValue,
                        Range = RangeAt(5, 1),
                        Formula = new TemplateFormulaDefinition { Formula = "=B5", FormulaR1C1 = "=RC[1]" }
                    },
                    new TemplateRegionDefinition
                    {
                        Role = TemplateRegionRole.RangeValue,
                        Range = RangeAt(2, 9),
                        Unit = "ppm"
                    }
                }
            };
            var service = new TemplateSaveService(
                new TemplateSyncClient(
                    new HttpClient(new StaticJsonHandler("{}")),
                    "http://localhost/api/templates"),
                repository);

            var result = service.Save(
                "Static Template",
                fingerprint,
                new[] { rule },
                null,
                true);
            var cachedRule = repository.FindByExactFingerprint(fingerprint.ExactFingerprint).Rules.Single();

            Assert.IsTrue(result.SavedToLocal);
            Assert.IsFalse(cachedRule.FixedStandardValue.HasValue);
            Assert.IsFalse(cachedRule.FixedReferenceRange.HasValue);
            Assert.IsNotNull(cachedRule.TemplateDefinition);
            Assert.AreEqual("ppm", cachedRule.TemplateDefinition.Regions.Single(region =>
                region.Role == TemplateRegionRole.RangeValue).Unit);
        }

        [TestMethod]
        public void TemplateSavePreservesManualStandardValuesWhenWorkbookHasStandardRange()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-manual-standard-local");
            var rule = Rule("manual standard", 100);
            rule.StandardValueSource = new ParameterSource { Name = "standard", Range = RangeAt(5, 1) };
            rule.ManualStandardValues = new List<ManualStandardValue>
            {
                new ManualStandardValue { PointIndex = 1, Value = 100 },
                new ManualStandardValue { PointIndex = 2, Value = 120 }
            };

            var service = new TemplateSaveService(
                new TemplateSyncClient(
                    new HttpClient(new StaticJsonHandler("{}")),
                    "http://localhost/api/excel-templates"),
                repository);
            service.Save("Manual Standard", fingerprint, new[] { rule }, null, true);

            var cachedRule = repository.FindByExactFingerprint(fingerprint.ExactFingerprint).Rules.Single();
            Assert.AreEqual(2, cachedRule.ManualStandardValues.Count);
            Assert.AreEqual(100d, cachedRule.FixedStandardValue.Value, 1e-12);

            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta> { new CellMeta { Row = 5, Column = 1, Text = "25" } }
                    }
                }
            };
            var resolved = new MeasurementRuleParameterResolver()
                .Apply(snapshot, new[] { cachedRule })
                .Single();
            Assert.AreEqual(100d, resolved.FixedStandardValue.Value, 1e-12);
        }

        [TestMethod]
        public void TemplateSaveDoesNotLetRemoteResponseDropManualStandardValues()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-manual-standard-remote");
            var submitted = Rule("manual remote", 100);
            submitted.StandardValueSource = new ParameterSource { Name = "standard", Range = RangeAt(5, 1) };
            submitted.ManualStandardValues = new List<ManualStandardValue>
            {
                new ManualStandardValue { PointIndex = 1, Value = 100 }
            };
            var remoteRule = Rule("manual remote", 25);
            remoteRule.StandardValueSource = new ParameterSource { Name = "standard", Range = RangeAt(5, 1) };
            var response = JsonConvert.SerializeObject(new
            {
                ok = true,
                data = new
                {
                    template_id = "manual-remote-1",
                    template_name = "Manual Remote",
                    version = 1,
                    fingerprint,
                    rules = new[] { remoteRule }
                }
            });
            var service = new TemplateSaveService(
                new TemplateSyncClient(
                    new HttpClient(new StaticJsonHandler(response)),
                    "http://localhost/api/excel-templates"),
                repository);

            service.Save("Manual Remote", fingerprint, new[] { submitted }, null, true);

            var cachedRule = repository.FindByExactFingerprint(fingerprint.ExactFingerprint).Rules.Single();
            Assert.AreEqual(1, cachedRule.ManualStandardValues.Count);
            Assert.AreEqual(100d, cachedRule.FixedStandardValue.Value, 1e-12);
        }

        [TestMethod]
        public void GenerationUsesPersistedManualStandardInsteadOfWorksheetValue()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "25" },
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 5, Column = 4, Formula = "=C5-A5" }
                        }
                    }
                }
            };
            var rule = Rule("manual generation", 100);
            rule.StandardValueSource = new ParameterSource { Name = "standard", Range = RangeAt(5, 1) };
            rule.ManualStandardValues = new List<ManualStandardValue>
            {
                new ManualStandardValue { PointIndex = 1, Value = 100 }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(31)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                new MeasurementRuleParameterResolver());

            var preview = useCase.Preview(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.Single() >= 99d);
            Assert.IsTrue(preview.RawValues.Single() <= 101d);
        }

        [TestMethod]
        public void TemplateSaveKeepsSubmittedFingerprintWhenRemoteResponseUsesAnotherFingerprint()
        {
            var repository = CreateRepository();
            var workbookFingerprint = Fingerprint("fp-after-calibration-item-delete");
            var remoteFingerprint = Fingerprint("remote-normalized-fingerprint");
            var response = JsonConvert.SerializeObject(new
            {
                data = new
                {
                    template_id = "remote-after-delete",
                    template_name = "Deleted calibration item",
                    version = 3,
                    fingerprint = remoteFingerprint,
                    rules = new[] { Rule("remaining calibration item", 10) }
                }
            });
            var service = new TemplateSaveService(
                new TemplateSyncClient(
                    new HttpClient(new StaticJsonHandler(response)),
                    "http://localhost/api/templates"),
                repository);

            service.Save(
                "Deleted calibration item",
                workbookFingerprint,
                new[] { Rule("remaining calibration item", 10) },
                null,
                createNew: false);

            var matched = repository.FindByExactFingerprint(workbookFingerprint.ExactFingerprint);
            Assert.IsNotNull(matched);
            Assert.AreEqual("remote-after-delete", matched.RemoteTemplateId);
            Assert.AreEqual(1, matched.Rules.Count);
            Assert.IsNull(repository.FindByExactFingerprint(remoteFingerprint.ExactFingerprint));
        }

        [TestMethod]
        public void TemplateDirectoryMetadataIsPersistedWithTemplateVariant()
        {
            var repository = CreateRepository();
            var metadata = new TemplateDirectoryMetadata
            {
                MeasurementDomain = "Gas",
                TemplateName = "Combustible alarm",
                VariantName = "Without repeatability",
                TemplateCode = "GAS-ALARM-002"
            };

            repository.SaveTemplate(
                "Combustible alarm + Without repeatability",
                Fingerprint("fp-directory-metadata"),
                new[] { Rule("remaining calibration item", 10) },
                directoryMetadata: metadata);

            var saved = repository.FindByExactFingerprint("fp-directory-metadata");
            Assert.IsNotNull(saved);
            Assert.AreEqual(metadata.MeasurementDomain, saved.DirectoryMetadata.MeasurementDomain);
            Assert.AreEqual(metadata.TemplateName, saved.DirectoryMetadata.TemplateName);
            Assert.AreEqual(metadata.VariantName, saved.DirectoryMetadata.VariantName);
            Assert.AreEqual(metadata.TemplateCode, saved.DirectoryMetadata.TemplateCode);
        }

        [TestMethod]
        public void TemplateSaveRejectsIncompleteRemoteResponseAndQueuesUpload()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-incomplete-response");
            var client = new TemplateSyncClient(
                new HttpClient(new StaticJsonHandler("{\"ok\":true}")),
                "http://localhost:3002/api/templates");
            var service = new TemplateSaveService(client, repository);

            var result = service.Save("Incomplete", fingerprint, new[] { Rule("rule", 10) }, null, false);

            Assert.IsFalse(result.SavedToRemote);
            Assert.AreEqual(TemplateSyncStatus.PendingUpload, result.LocalSyncStatus);
            Assert.AreEqual(TemplateSyncStatus.PendingUpload, repository.FindByExactFingerprint(fingerprint.ExactFingerprint).LocalSyncStatus);
        }

        [TestMethod]
        public void TemplateSaveSendsOverwriteIdentityToRemote()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-overwrite-contract");
            repository.SaveTemplate("Existing", fingerprint, new[] { Rule("existing", 10) }, null);
            var response = JsonConvert.SerializeObject(new
            {
                data = new
                {
                    template_id = fingerprint.ExactFingerprint,
                    template_name = "Existing",
                    version = 2,
                    fingerprint,
                    rules = new[] { Rule("accepted", 10) }
                }
            });
            var handler = new StaticJsonHandler(response);
            var service = new TemplateSaveService(
                new TemplateSyncClient(new HttpClient(handler), "http://localhost/api/templates"),
                repository);

            service.Save("Existing", fingerprint, new[] { Rule("updated", 10) }, null, false);

            StringAssert.Contains(handler.LastRequestBody, "\"templateId\":\"fp-overwrite-contract\"");
            StringAssert.Contains(handler.LastRequestBody, "\"createNew\":false");
        }

        [TestMethod]
        public void RowMappingBuilderAssignsMergedRequirementToCoveredRows()
        {
            var requirementRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 5, EndColumn = 5 };
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "10" },
                            new CellMeta { Row = 6, Column = 1, Text = "20" },
                            new CellMeta { Row = 5, Column = 4, Formula = "=C5-A5" },
                            new CellMeta { Row = 6, Column = 4, Formula = "=C6-A6" },
                            new CellMeta { Row = 5, Column = 5, Text = "±0.5", IsMerged = true, MergeRange = requirementRange },
                            new CellMeta { Row = 6, Column = 5, IsMerged = true, MergeRange = requirementRange }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 2, EndColumn = 3 },
                StandardValueSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 1, EndColumn = 1 } },
                ErrorSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 6, StartColumn = 4, EndColumn = 4 } },
                MpeSource = new ParameterSource { Range = requirementRange },
                WritableCells = new List<CellAddress>
                {
                    new CellAddress { Row = 5, Column = 2 }, new CellAddress { Row = 5, Column = 3 },
                    new CellAddress { Row = 6, Column = 2 }, new CellAddress { Row = 6, Column = 3 }
                }
            };

            var mappings = new RowMappingBuilder().Build(snapshot, rule);

            Assert.AreEqual(2, mappings.Count);
            Assert.IsTrue(mappings.All(item => item.IsComplete));
            Assert.IsTrue(mappings.All(item => item.TechnicalRequirementRange.StartRow == 5 && item.TechnicalRequirementRange.EndRow == 6));
        }

        [TestMethod]
        public void PendingUploadWritesRemoteReturnedTemplateToLocalCache()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-pending-remote-save");
            repository.SaveTemplate(
                "Pending Draft",
                fingerprint,
                new[] { Rule("pending local", 10) },
                null,
                localSyncStatus: TemplateSyncStatus.PendingUpload);
            var remoteRule = Rule("pending remote accepted", 88);
            var response = JsonConvert.SerializeObject(new
            {
                ok = true,
                data = new
                {
                    template_id = "remote-pending-1",
                    template_name = "Pending Remote Accepted",
                    version = 3,
                    updated_at = DateTime.UtcNow.ToString("o"),
                    status = "enabled",
                    fingerprint,
                    rules = new[] { remoteRule }
                }
            });
            var client = new TemplateSyncClient(
                new HttpClient(new StaticJsonHandler(response)),
                "http://localhost:3002/api/templates");
            var service = new TemplatePendingUploadService(client, repository);

            var uploaded = service.UploadPendingTemplates();

            var cached = repository.FindByExactFingerprint(fingerprint.ExactFingerprint);
            Assert.AreEqual(1, uploaded);
            Assert.AreEqual(TemplateSyncStatus.Synced, cached.LocalSyncStatus);
            Assert.AreEqual("remote-pending-1", cached.RemoteTemplateId);
            Assert.AreEqual("Pending Remote Accepted", cached.TemplateName);
            Assert.AreEqual(88, cached.Rules[0].FixedStandardValue);
        }

        [TestMethod]
        public void FullSyncPendingUploadWritesRemoteReturnedTemplateToLocalCache()
        {
            var repository = CreateRepository();
            var fingerprint = Fingerprint("fp-full-sync-pending");
            repository.SaveTemplate(
                "Full Sync Pending",
                fingerprint,
                new[] { Rule("full sync local", 10) },
                null,
                localSyncStatus: TemplateSyncStatus.PendingUpload);
            var remoteRule = Rule("full sync accepted", 77);
            var saveResponse = JsonConvert.SerializeObject(new
            {
                ok = true,
                data = new
                {
                    template_id = "remote-full-sync-1",
                    template_name = "Full Sync Remote Accepted",
                    version = 4,
                    updated_at = DateTime.UtcNow.ToString("o"),
                    status = "enabled",
                    fingerprint,
                    rules = new[] { remoteRule }
                }
            });
            var listResponse = JsonConvert.SerializeObject(new { data = new object[0] });
            var client = new TemplateSyncClient(
                new HttpClient(new SequenceJsonHandler(saveResponse, listResponse)),
                "http://localhost:3002/api/templates");
            var useCase = new SyncTemplateUseCase(client, repository);

            var result = useCase.SyncAsync().GetAwaiter().GetResult();

            var cached = repository.FindByExactFingerprint(fingerprint.ExactFingerprint);
            Assert.IsTrue(result.Succeeded);
            Assert.AreEqual(1, result.PendingUploadsSucceeded);
            Assert.AreEqual(TemplateSyncStatus.Synced, cached.LocalSyncStatus);
            Assert.AreEqual("remote-full-sync-1", cached.RemoteTemplateId);
            Assert.AreEqual("Full Sync Remote Accepted", cached.TemplateName);
            Assert.AreEqual(77, cached.Rules[0].FixedStandardValue);
        }

        [TestMethod]
        public void FormulaVerifierAcceptsExcelPercentageDisplayResult()
        {
            var snapshot = new WorkbookSnapshot
            {
                WorkbookName = "book.xlsx",
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 4, Text = "0.30%", Formula = "=(C5-A5)/A5" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "relative",
                ErrorType = ErrorType.Relative,
                FixedStandardValue = 100,
                FixedMpe = 0.005,
                ErrorSource = new ParameterSource { Range = Range("D5:D5") },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Scale = ErrorFormulaScale.RelativeToStandardValue,
                    FormulaMultipliesBy100 = false
                }
            };

            new FormulaResultVerifier().Verify(snapshot, new[] { rule });
        }

        [TestMethod]
        public void FormulaVerifierRejectsRepeatabilityWithIdenticalWrittenMeasurements()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = Enumerable.Range(3, 6)
                            .Select(column => new CellMeta
                            {
                                Row = 5,
                                Column = column,
                                Text = "37",
                                RawValueText = "37",
                                NumberFormat = "0"
                            })
                            .Concat(new[]
                            {
                                new CellMeta
                                {
                                    Row = 5,
                                    Column = 9,
                                    Text = "0.0",
                                    RawValueText = "0",
                                    NumberFormat = "0.0",
                                    Formula = "=STDEV.S(C5:H5)/AVERAGE(C5:H5)*100"
                                }
                            })
                            .ToList()
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "重复性",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 3, EndColumn = 8 },
                ErrorSource = new ParameterSource { Range = RangeAt(5, 9) },
                FixedMpe = 2,
                ErrorFormula = new ErrorFormulaInfo { HasFormula = true }
            };

            var exception = Assert.ThrowsException<RepeatabilityVerificationException>(() =>
                new FormulaResultVerifier().Verify(snapshot, new[] { rule }));

            StringAssert.Contains(exception.Message, "重复性为 0");
        }

        [TestMethod]
        public void FormulaVerifierRejectsRepeatabilityRoundedToZeroByErrorPrecision()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 3, Text = "37", RawValueText = "37", NumberFormat = "0" },
                            new CellMeta { Row = 5, Column = 4, Text = "38", RawValueText = "38", NumberFormat = "0" },
                            new CellMeta { Row = 5, Column = 5, Text = "37", RawValueText = "37", NumberFormat = "0" },
                            new CellMeta { Row = 5, Column = 6, Text = "38", RawValueText = "38", NumberFormat = "0" },
                            new CellMeta { Row = 5, Column = 7, Text = "37", RawValueText = "37", NumberFormat = "0" },
                            new CellMeta { Row = 5, Column = 8, Text = "38", RawValueText = "38", NumberFormat = "0" },
                            new CellMeta
                            {
                                Row = 5,
                                Column = 9,
                                Text = "0.0",
                                RawValueText = "0.04",
                                NumberFormat = "0.0",
                                Formula = "=STDEV.S(C5:H5)/AVERAGE(C5:H5)*100"
                            }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "重复性",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 3, EndColumn = 8 },
                ErrorSource = new ParameterSource { Range = RangeAt(5, 9) },
                FixedMpe = 2,
                ErrorFormula = new ErrorFormulaInfo { HasFormula = true }
            };

            var exception = Assert.ThrowsException<RepeatabilityVerificationException>(() =>
                new FormulaResultVerifier().Verify(snapshot, new[] { rule }));

            StringAssert.Contains(exception.Message, "误差列分辨力下显示为 0");
        }

        [TestMethod]
        public void FormulaVerifierRejectsAnyErrorDisplayedAsZeroAtErrorCellPrecision()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta
                            {
                                Row = 5,
                                Column = 4,
                                Text = "0.0",
                                RawValueText = "0.04",
                                NumberFormat = "0.0",
                                Formula = "=C5-A5"
                            }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                ErrorType = ErrorType.Absolute,
                FixedMpe = 1,
                ErrorSource = new ParameterSource { Range = RangeAt(5, 4) },
                ErrorFormula = new ErrorFormulaInfo { HasFormula = true }
            };

            var exception = Assert.ThrowsException<DisplayedErrorZeroVerificationException>(() =>
                new FormulaResultVerifier().Verify(snapshot, new[] { rule }));

            StringAssert.Contains(exception.Message, "当前分辨力下显示为 0");
        }

        [TestMethod]
        public void FormulaVerifierAcceptsFloatingPointRoundingAtInclusiveLimit()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta
                            {
                                Row = 5,
                                Column = 4,
                                RawValueText = "5.000000000000001",
                                Text = "5.0",
                                Formula = "=C5-A5"
                            }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                ErrorSource = new ParameterSource { Range = Range("D5:D5") },
                FixedMpe = 5,
                RequirementOperator = TechnicalRequirementOperator.PlusMinus,
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Scale = ErrorFormulaScale.Absolute
                }
            };

            new FormulaResultVerifier().Verify(snapshot, new[] { rule });
        }

        [TestMethod]
        public void FormulaVerifierUsesRawValueForConditionalAbsoluteBranch()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 6, Column = 26, Text = "10" },
                            new CellMeta { Row = 7, Column = 18, Text = "绝对误差（µmol/mol）" },
                            new CellMeta { Row = 9, Column = 2, Text = "2.0", RawValueText = "2" },
                            new CellMeta { Row = 9, Column = 15, Text = "3.5", RawValueText = "3.5", Formula = "=AVERAGE(F9:N9)" },
                            new CellMeta
                            {
                                Row = 9,
                                Column = 18,
                                Text = "150%",
                                DisplayText = "150%",
                                RawValueText = "1.5",
                                Formula = "=IF($Z$6<=10,O9-B9,(O9-B9)/B9*100)",
                                NumberFormat = "0%"
                            },
                            new CellMeta
                            {
                                Row = 9,
                                Column = 22,
                                Text = "±2µmol/mol",
                                Formula = "=IF($Z$6<=10,\"±2µmol/mol\",\"±10%\")"
                            }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 9, EndRow = 9, StartColumn = 6, EndColumn = 14 },
                StandardValueSource = new ParameterSource { Range = RangeAt(9, 2) },
                AverageSource = new ParameterSource { Range = RangeAt(9, 15) },
                ErrorSource = new ParameterSource { Range = RangeAt(9, 18) },
                MpeSource = new ParameterSource { Range = RangeAt(9, 22) },
                RangeSource = new ParameterSource { Range = RangeAt(6, 26) },
                FixedMpe = 2,
                ErrorFormula = new ErrorFormulaInfo { HasFormula = true }
            };

            new FormulaResultVerifier().Verify(snapshot, new[] { rule });

            Assert.AreEqual(ErrorFormulaScale.Absolute, rule.ErrorFormula.Scale);
        }

        [TestMethod]
        public void FormulaVerifierUsesRawValueForConditionalRelativeBranch()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 6, Column = 26, Text = "100" },
                            new CellMeta { Row = 9, Column = 2, Text = "20", RawValueText = "20" },
                            new CellMeta { Row = 9, Column = 15, Text = "21.6", RawValueText = "21.6", Formula = "=AVERAGE(F9:N9)" },
                            new CellMeta
                            {
                                Row = 9,
                                Column = 18,
                                Text = "800%",
                                DisplayText = "800%",
                                RawValueText = "8",
                                Formula = "=IF($Z$6<=10,O9-B9,(O9-B9)/B9*100)",
                                NumberFormat = "0%"
                            },
                            new CellMeta
                            {
                                Row = 9,
                                Column = 22,
                                Text = "±10%",
                                Formula = "=IF($Z$6<=10,\"±2µmol/mol\",\"±10%\")"
                            }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 9, EndRow = 9, StartColumn = 6, EndColumn = 14 },
                StandardValueSource = new ParameterSource { Range = RangeAt(9, 2) },
                AverageSource = new ParameterSource { Range = RangeAt(9, 15) },
                ErrorSource = new ParameterSource { Range = RangeAt(9, 18) },
                MpeSource = new ParameterSource { Range = RangeAt(9, 22) },
                RangeSource = new ParameterSource { Range = RangeAt(6, 26) },
                FixedMpe = 0.1,
                ErrorFormula = new ErrorFormulaInfo { HasFormula = true }
            };

            new FormulaResultVerifier().Verify(snapshot, new[] { rule });

            Assert.AreEqual(ErrorFormulaScale.RelativeToStandardValue, rule.ErrorFormula.Scale);
        }

        [TestMethod]
        public void StructureAnalyzerRecordsSupplementalFormulas()
        {
            var snapshot = new WorkbookSnapshot
            {
                WorkbookName = "book.xlsx",
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 3, Column = 1, Text = "10" },
                            new CellMeta { Row = 3, Column = 3, Text = "10.2" },
                            new CellMeta { Row = 3, Column = 4, Text = "0.2", Formula = "=C3-A3" },
                            new CellMeta { Row = 3, Column = 5, Text = "±0.5", Formula = "=IF(A3<=10,\"±0.5\",\"±1\")" },
                            new CellMeta { Row = 3, Column = 6, Text = "合格", Formula = "=IF(ABS(D3)<=0.5,\"合格\",\"不合格\")" },
                            new CellMeta { Row = 3, Column = 7, Text = "0.1", Formula = "=D3/2" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = RangeAt(3, 3),
                StandardValueSource = new ParameterSource { Range = RangeAt(3, 1) },
                ErrorSource = new ParameterSource { Range = RangeAt(3, 4) },
                MpeSource = new ParameterSource { Range = RangeAt(3, 5) },
                ResultSource = new ParameterSource { Range = RangeAt(3, 6) },
                UncertaintySource = new ParameterSource { Range = RangeAt(3, 7) },
                FixedMpe = 0.5
            };

            new MeasurementRuleStructureAnalyzer().Apply(snapshot, new[] { rule });

            Assert.IsTrue(rule.ErrorFormula.HasFormula);
            Assert.AreEqual("=C3-A3", rule.ErrorFormula.Formula);
            Assert.IsTrue(rule.ErrorFormula.TechnicalRequirementFormulaResolved);
            Assert.AreEqual("=IF(A3<=10,\"±0.5\",\"±1\")", rule.ErrorFormula.TechnicalRequirementFormula);
            Assert.IsTrue(rule.ErrorFormula.ResultFormulaResolved);
            Assert.AreEqual("=IF(ABS(D3)<=0.5,\"合格\",\"不合格\")", rule.ErrorFormula.ResultFormula);
            Assert.IsTrue(rule.ErrorFormula.UncertaintyFormulaResolved);
            Assert.AreEqual("=D3/2", rule.ErrorFormula.UncertaintyFormula);
        }

        [TestMethod]
        public void StructureAnalyzerUsesDisplayedAbsoluteBranchForConditionalErrorFormula()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 6, Column = 26, Text = "10" },
                            new CellMeta { Row = 6, Column = 21, Text = "量程：" },
                            new CellMeta { Row = 7, Column = 18, Text = "绝对误差（µmol/mol）" },
                            new CellMeta { Row = 9, Column = 2, Text = "2.0" },
                            new CellMeta { Row = 9, Column = 15, Text = "2.0", Formula = "=AVERAGE(F9:N9)" },
                            new CellMeta
                            {
                                Row = 9,
                                Column = 18,
                                Text = "0.0",
                                Formula = "=IF($Z$6<=10,O9-B9,(O9-B9)/B9*100)"
                            },
                            new CellMeta { Row = 9, Column = 22, Text = "±2µmol/mol" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 9, EndRow = 9, StartColumn = 6, EndColumn = 14 },
                StandardValueSource = new ParameterSource { Range = RangeAt(9, 2) },
                AverageSource = new ParameterSource { Range = RangeAt(9, 15) },
                ErrorSource = new ParameterSource { Range = RangeAt(9, 18) },
                MpeSource = new ParameterSource { Range = RangeAt(9, 22) },
                RangeSource = new ParameterSource { Range = RangeAt(6, 26) }
            };

            new MeasurementRuleStructureAnalyzer().Apply(snapshot, new[] { rule });

            Assert.AreEqual(ErrorFormulaScale.Absolute, rule.ErrorFormula.Scale);
            Assert.IsFalse(rule.ErrorFormula.FormulaDividesByReferenceRange);
            Assert.AreEqual(ErrorType.Absolute, GenerationRuleValidator.ResolveGenerationErrorType(rule));
        }

        [TestMethod]
        public void CachedRuleRefreshesConditionalFormulaScaleBeforeGeneration()
        {
            var mergedRequirementRange = new CellRange
            {
                SheetName = "Sheet1",
                StartRow = 5,
                EndRow = 7,
                StartColumn = 5,
                EndColumn = 5
            };
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 2, Column = 6, Text = "10" },
                            new CellMeta { Row = 3, Column = 4, Text = "绝对误差（µmol/mol）" },
                            new CellMeta { Row = 5, Column = 1, Text = "2.0" },
                            new CellMeta { Row = 6, Column = 1, Text = "5.0" },
                            new CellMeta { Row = 7, Column = 1, Text = "8.0" },
                            new CellMeta { Row = 5, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 6, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 7, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 5, Column = 4, Text = "0.0", Formula = "=IF($F$2<=10,C5-A5,(C5-A5)/A5*100)" },
                            new CellMeta { Row = 6, Column = 4, Text = "0.0", Formula = "=IF($F$2<=10,C6-A6,(C6-A6)/A6*100)" },
                            new CellMeta { Row = 7, Column = 4, Text = "0.0", Formula = "=IF($F$2<=10,C7-A7,(C7-A7)/A7*100)" },
                            new CellMeta { Row = 5, Column = 5, Text = "±2µmol/mol", IsMerged = true, MergeRange = mergedRequirementRange },
                            new CellMeta { Row = 6, Column = 5, IsMerged = true, MergeRange = mergedRequirementRange },
                            new CellMeta { Row = 7, Column = 5, IsMerged = true, MergeRange = mergedRequirementRange }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 3, EndColumn = 3 },
                StandardValueSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 1, EndColumn = 1 } },
                ErrorSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 4, EndColumn = 4 } },
                MpeSource = new ParameterSource { Range = mergedRequirementRange, ValuePattern = "mpe:absolute:scale=1:op=plusminus" },
                RangeSource = new ParameterSource { Range = RangeAt(2, 6) },
                FixedStandardValue = 2,
                FixedMpe = 2,
                FixedReferenceRange = 10,
                NegativeDirectionOnly = true,
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=IF($F$2<=10,C5-A5,(C5-A5)/A5*100)",
                    ReferencesMeasurement = true,
                    Scale = ErrorFormulaScale.RelativeToReferenceRange,
                    FormulaMultipliesBy100 = true,
                    FormulaDividesByReferenceRange = true,
                    TechnicalRequirementFormula = "=E5",
                    ResultFormula = "=G5"
                }
            };
            var configuration = new GenerationConfiguration();
            var useCase = new GenerateMeasurementUseCase(
                current => new MeasurementValueGenerator(current, new Random(17)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                new MeasurementRuleParameterResolver());

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.AreEqual(ErrorFormulaScale.Absolute, preview.Rule.ErrorFormula.Scale);
            Assert.AreEqual(3, preview.RawValues.Count);
            var standards = new[] { 2d, 5d, 8d };
            Assert.IsTrue(preview.RawValues.Select((value, index) => value >= 0 && Math.Abs(value - standards[index]) <= 2).All(value => value));
        }

        [TestMethod]
        public void ParameterResolverReplacesCachedAbsolutePatternForRelativeConditionalBranch()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 6, Column = 26, Text = "100", RawValueText = "100" },
                            new CellMeta { Row = 9, Column = 2, Text = "20.0", RawValueText = "20" },
                            new CellMeta
                            {
                                Row = 9,
                                Column = 22,
                                Text = "±10%",
                                DisplayText = "±10%",
                                RawValueText = "±10%",
                                Formula = "=IF($Z$6<=10,\"±2µmol/mol\",\"±10%\")"
                            }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = RangeAt(9, 6),
                StandardValueSource = new ParameterSource { Range = RangeAt(9, 2) },
                MpeSource = new ParameterSource
                {
                    Range = RangeAt(9, 22),
                    ValuePattern = "mpe:absolute:scale=1:op=plusminus"
                },
                RangeSource = new ParameterSource { Range = RangeAt(6, 26) }
            };

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(ErrorType.Relative, resolved.ErrorType);
            Assert.AreEqual(0.1, resolved.FixedMpe.Value, 1e-12);
            StringAssert.StartsWith(resolved.MpeSource.ValuePattern, "mpe:relative:scale=0.01");
        }

        [TestMethod]
        public void ParameterResolverReplacesCachedRelativePatternForAbsoluteConditionalBranch()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 6, Column = 26, Text = "10", RawValueText = "10" },
                            new CellMeta { Row = 9, Column = 2, Text = "2.0", RawValueText = "2" },
                            new CellMeta
                            {
                                Row = 9,
                                Column = 22,
                                Text = "±2µmol/mol",
                                DisplayText = "±2µmol/mol",
                                RawValueText = "±2µmol/mol",
                                Formula = "=IF($Z$6<=10,\"±2µmol/mol\",\"±10%\")"
                            }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = RangeAt(9, 6),
                StandardValueSource = new ParameterSource { Range = RangeAt(9, 2) },
                MpeSource = new ParameterSource
                {
                    Range = RangeAt(9, 22),
                    ValuePattern = "mpe:relative:scale=0.01:op=plusminus"
                },
                RangeSource = new ParameterSource { Range = RangeAt(6, 26) }
            };

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(ErrorType.Absolute, resolved.ErrorType);
            Assert.AreEqual(2, resolved.FixedMpe.Value, 1e-12);
            StringAssert.StartsWith(resolved.MpeSource.ValuePattern, "mpe:absolute:scale=1");
        }

        [TestMethod]
        public void ParameterResolverPreservesReferencedPatternWhenValueOmitsHeaderUnit()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 9, Column = 20, Text = "技术要求（%FS）" },
                            new CellMeta
                            {
                                Row = 11,
                                Column = 20,
                                Text = "±5",
                                DisplayText = "±5",
                                RawValueText = "±5"
                            },
                            new CellMeta { Row = 8, Column = 22, Text = "0", RawValueText = "0" },
                            new CellMeta { Row = 8, Column = 27, Text = "100", RawValueText = "100" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = RangeAt(11, 6),
                MpeSource = new ParameterSource
                {
                    Range = RangeAt(11, 20),
                    ValuePattern = "mpe:referenced:scale=0.01:op=plusminus"
                },
                RangeSource = new ParameterSource
                {
                    Range = new CellRange
                    {
                        SheetName = "Sheet1",
                        StartRow = 8,
                        EndRow = 8,
                        StartColumn = 22,
                        EndColumn = 27
                    }
                }
            };

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(ErrorType.Referenced, resolved.ErrorType);
            Assert.AreEqual(0.05, resolved.FixedMpe.Value, 1e-12);
            Assert.AreEqual(100, resolved.FixedReferenceRange.Value, 1e-12);
            StringAssert.StartsWith(resolved.MpeSource.ValuePattern, "mpe:referenced:scale=0.01");
        }

        [TestMethod]
        public void StructureAnalyzerRecognizesRangeOnlyWhenItIsInFormulaDenominator()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 8, Column = 22, Text = "0" },
                            new CellMeta { Row = 8, Column = 27, Text = "100" },
                            new CellMeta { Row = 11, Column = 21, Text = "0.0%", Formula = "=(R11-E11)/($AA$8-$V$8)*100" },
                            new CellMeta { Row = 11, Column = 24, Text = "±5%FS" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "引用误差",
                TargetRange = RangeAt(11, 18),
                StandardValueSource = new ParameterSource { Range = RangeAt(11, 5) },
                ErrorSource = new ParameterSource { Range = RangeAt(11, 21) },
                MpeSource = new ParameterSource { Range = RangeAt(11, 24) },
                RangeSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 8, EndRow = 8, StartColumn = 22, EndColumn = 27 }
                }
            };

            new MeasurementRuleStructureAnalyzer().Apply(snapshot, new[] { rule });

            Assert.AreEqual(ErrorFormulaScale.RelativeToReferenceRange, rule.ErrorFormula.Scale);
            Assert.IsTrue(rule.ErrorFormula.FormulaDividesByReferenceRange);
        }

        [TestMethod]
        public void ParameterResolverUsesEndpointMagnitudeForSymmetricBipolarReferenceRange()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 1, Text = "7.00" },
                            new CellMeta { Row = 5, Column = 2, Text = "±0.1%FS" },
                            new CellMeta { Row = 5, Column = 3, Text = "-1999~1999" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "PH",
                TargetRange = RangeAt(5, 4),
                StandardValueSource = new ParameterSource { Range = RangeAt(5, 1) },
                MpeSource = new ParameterSource { Range = RangeAt(5, 2) },
                RangeSource = new ParameterSource { Range = RangeAt(5, 3) }
            };

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(ErrorType.Referenced, resolved.ErrorType);
            Assert.AreEqual(0.001, resolved.FixedMpe.Value, 1e-12);
            Assert.AreEqual(1999, resolved.FixedReferenceRange.Value, 1e-12);
        }

        [TestMethod]
        public void StructureAnalyzerAlignsStandardValueSourceToErrorFormula()
        {
            var sheet = new SheetSnapshot
            {
                Name = "Sheet1",
                Cells = new List<CellMeta>
                {
                    new CellMeta
                    {
                        Row = 11,
                        Column = 2,
                        Text = "3.00",
                        MergeRange = new CellRange { SheetName = "Sheet1", StartRow = 11, EndRow = 13, StartColumn = 2, EndColumn = 4 }
                    },
                    new CellMeta
                    {
                        Row = 11,
                        Column = 5,
                        Text = "10.0",
                        MergeRange = new CellRange { SheetName = "Sheet1", StartRow = 11, EndRow = 11, StartColumn = 5, EndColumn = 8 }
                    },
                    new CellMeta
                    {
                        Row = 12,
                        Column = 5,
                        Text = "40.0",
                        MergeRange = new CellRange { SheetName = "Sheet1", StartRow = 12, EndRow = 12, StartColumn = 5, EndColumn = 8 }
                    },
                    new CellMeta
                    {
                        Row = 13,
                        Column = 5,
                        Text = "60.0",
                        MergeRange = new CellRange { SheetName = "Sheet1", StartRow = 13, EndRow = 13, StartColumn = 5, EndColumn = 8 }
                    }
                }
            };
            var snapshot = new WorkbookSnapshot { Sheets = new List<SheetSnapshot> { sheet } };
            var rule = new MeasurementRule
            {
                FieldName = "indication error",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 11, EndRow = 13, StartColumn = 9, EndColumn = 17 },
                StandardValueSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 11, EndRow = 13, StartColumn = 2, EndColumn = 4 }
                },
                AverageSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 11, EndRow = 13, StartColumn = 18, EndColumn = 20 }
                },
                ErrorSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 11, EndRow = 13, StartColumn = 21, EndColumn = 23 }
                },
                RangeSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 8, EndRow = 8, StartColumn = 22, EndColumn = 27 }
                },
                FixedMpe = 0.05,
                FixedReferenceRange = 100,
                WritableCells = new List<CellAddress>
                {
                    new CellAddress { Row = 11, Column = 9 },
                    new CellAddress { Row = 12, Column = 9 },
                    new CellAddress { Row = 13, Column = 9 }
                },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=(R11-E11)/($AA$8-$V$8)*100",
                    ReferencesAverage = true,
                    AverageFormulaResolved = true,
                    AverageFormula = "=AVERAGE(I11:Q11)",
                    TechnicalRequirementFormula = "=X11",
                    ResultFormula = "=AE11",
                    Scale = ErrorFormulaScale.RelativeToReferenceRange,
                    FormulaMultipliesBy100 = true,
                    FormulaDividesByReferenceRange = true
                }
            };

            new MeasurementRuleStructureAnalyzer().Apply(snapshot, new[] { rule });
            var mappings = new RowMappingBuilder().Build(snapshot, rule);

            Assert.AreEqual(5, rule.StandardValueSource.Range.StartColumn);
            Assert.AreEqual(8, rule.StandardValueSource.Range.EndColumn);
            Assert.AreEqual(11, rule.StandardValueSource.Range.StartRow);
            Assert.AreEqual(13, rule.StandardValueSource.Range.EndRow);
            Assert.IsTrue(rule.ErrorFormula.ReferencesStandardValue);
            Assert.AreEqual(3, mappings.Count);
            Assert.AreEqual(11, mappings[0].StandardValueRange.StartRow);
            Assert.AreEqual(12, mappings[1].StandardValueRange.StartRow);
            Assert.AreEqual(13, mappings[2].StandardValueRange.StartRow);
        }

        [TestMethod]
        public void ManualStandardValuesApplyToTheirConfiguredMeasurementPoints()
        {
            var snapshot = new WorkbookSnapshot
            {
                WorkbookName = "book.xlsx",
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 5, Column = 3, Text = string.Empty },
                            new CellMeta { Row = 6, Column = 3, Text = string.Empty },
                            new CellMeta { Row = 7, Column = 3, Text = string.Empty }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "manual multi-point",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 7, StartColumn = 3, EndColumn = 3 },
                FixedStandardValue = 10,
                FixedMpe = 0.5,
                FormatRule = new FormatRule { DecimalPlaces = 3 },
                ManualStandardValues = new List<ManualStandardValue>
                {
                    new ManualStandardValue { PointIndex = 1, Value = 10 },
                    new ManualStandardValue { PointIndex = 3, Value = 30 }
                }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(23)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.AreEqual(2, preview.WritableCells.Count);
            CollectionAssert.AreEqual(new[] { 5, 7 }, preview.WritableCells.Select(cell => cell.Row).ToArray());
            Assert.AreEqual(2, preview.RawValues.Count);
            Assert.IsTrue(Math.Abs(preview.RawValues[0] - 10) <= 0.5);
            Assert.IsTrue(Math.Abs(preview.RawValues[1] - 30) <= 0.5);
        }

        [TestMethod]
        public void LocalTemplateFuzzyMatchIsCandidateOnly()
        {
            var repository = CreateRepository();
            var saved = Fingerprint("fp-saved");
            saved.FuzzyFingerprint = "same-layout";
            repository.SaveTemplate("Candidate", saved, new[] { Rule("rule", 10) });

            var current = Fingerprint("fp-current");
            current.FuzzyFingerprint = "same-layout";
            var candidate = repository.FindBestMatch(current);

            Assert.IsNotNull(candidate);
            Assert.AreEqual(75, candidate.MatchScore);
            Assert.AreNotEqual(current.ExactFingerprint, candidate.ExactFingerprint);
        }

        [TestMethod]
        public void LegacyTemplateFingerprintMatchesWhenStableMetadataAndLayoutMatch()
        {
            var repository = CreateRepository();
            var saved = Fingerprint("fp-saved");
            saved.FuzzyFingerprint = "same-layout";
            saved.StructureSignature = string.Empty;
            repository.SaveTemplate("Legacy", saved, new[] { Rule("rule", 10) });

            var current = Fingerprint("fp-current");
            current.FuzzyFingerprint = "same-layout";
            current.StructureSignature = "new-structure-signature";

            var matched = repository.FindByExactOrLegacyCompatibleFingerprint(current);

            Assert.IsNotNull(matched);
            Assert.AreEqual("fp-saved", matched.ExactFingerprint);
            Assert.AreEqual(100, matched.MatchScore);
            Assert.AreEqual("legacy compatible fingerprint", matched.MatchReason);
        }

        [TestMethod]
        public void StoredTemplateMatchesByStructureAfterHeaderStandardValueNormalization()
        {
            var repository = CreateRepository();
            var saved = Fingerprint("old-header-value-fingerprint");
            saved.StructureSignature = "stable-structure";
            saved.HeaderTexts.Add("10 ppm");
            repository.SaveTemplate("Existing", saved, new[] { Rule("rule", 10) });

            var current = Fingerprint("normalized-header-fingerprint");
            current.StructureSignature = "stable-structure";

            var matched = repository.FindByExactOrLegacyCompatibleFingerprint(current);

            Assert.IsNotNull(matched);
            Assert.AreEqual("old-header-value-fingerprint", matched.ExactFingerprint);
            Assert.AreEqual(100, matched.MatchScore);
            Assert.AreEqual("structure signature", matched.MatchReason);
        }

        [TestMethod]
        public void LegacyTemplateFingerprintDoesNotMatchWhenHeadersDiffer()
        {
            var repository = CreateRepository();
            var saved = Fingerprint("fp-saved");
            saved.FuzzyFingerprint = "same-layout";
            repository.SaveTemplate("Legacy", saved, new[] { Rule("rule", 10) });

            var current = Fingerprint("fp-current");
            current.FuzzyFingerprint = "same-layout";
            current.StructureSignature = "new-structure-signature";
            current.HeaderTexts = new List<string> { "Different header" };

            Assert.IsNull(repository.FindByExactOrLegacyCompatibleFingerprint(current));
        }

        [TestMethod]
        public void SavingIncompleteNumericRuleIsRejected()
        {
            var repository = CreateRepository();
            var incompleteRule = new MeasurementRule
            {
                FieldName = "numeric",
                IsEnabled = true,
                TargetRange = Range("C5:C5"),
                FixedStandardValue = 10,
                FixedMpe = 0.5
            };

            var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                repository.SaveTemplate("Incomplete", Fingerprint("fp-incomplete"), new[] { incompleteRule }));

            StringAssert.Contains(exception.Message, "误差");
        }

        private static MeasurementRule SharedStandardRule(string name, int row)
        {
            return new MeasurementRule
            {
                FieldName = name,
                TargetRange = new CellRange
                {
                    SheetName = "Sheet1",
                    StartRow = row,
                    EndRow = row,
                    StartColumn = 3,
                    EndColumn = 4
                },
                FixedStandardValue = 10,
                FixedMpe = 1,
                FormatRule = new FormatRule { DecimalPlaces = 3, UnitSuffix = "V" }
            };
        }

        private static WorkbookSnapshot BuildTemplateDefinitionSnapshot()
        {
            return new WorkbookSnapshot
            {
                WorkbookName = "template.xlsx",
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 1, Column = 1, Text = "\u793A\u503C\u8BEF\u5DEE" },
                            new CellMeta { Row = 2, Column = 1, Text = "\u6807\u51C6\u503C(ppm)" },
                            new CellMeta { Row = 2, Column = 3, Text = "\u6D4B\u91CF\u503C" },
                            new CellMeta { Row = 2, Column = 5, Text = "\u5E73\u5747\u503C" },
                            new CellMeta { Row = 2, Column = 6, Text = "\u8BEF\u5DEE(%FS)" },
                            new CellMeta { Row = 2, Column = 7, Text = "\u7B26\u53F7" },
                            new CellMeta { Row = 2, Column = 8, Text = "\u6280\u672F\u8981\u6C42(%FS)" },
                            new CellMeta { Row = 2, Column = 9, Text = "\u91CF\u7A0B(ppm)" },
                            new CellMeta { Row = 2, Column = 10, Text = "\u7ED3\u679C" },
                            new CellMeta
                            {
                                Row = 4,
                                Column = 1,
                                Text = "5",
                                RawValueText = "5",
                                Formula = "=B4*0.5",
                                FormulaR1C1 = "=RC[1]*0.5",
                                NumberFormat = "0.0"
                            },
                            new CellMeta { Row = 4, Column = 3, NumberFormat = "0.00" },
                            new CellMeta { Row = 4, Column = 4, NumberFormat = "0.00" },
                            new CellMeta
                            {
                                Row = 4,
                                Column = 5,
                                Text = "5.1",
                                Formula = "=AVERAGE(C4:D4)",
                                FormulaR1C1 = "=AVERAGE(RC[-2]:RC[-1])",
                                NumberFormat = "0.00"
                            },
                            new CellMeta
                            {
                                Row = 4,
                                Column = 6,
                                Text = "1.0",
                                Formula = "=IF($I$4>0,(E4-A4)/$I$4*100,(E4-A4)/A4*100)",
                                FormulaR1C1 = "=IF(RC[3]>0,(RC[-1]-RC[-5])/RC[3]*100,(RC[-1]-RC[-5])/RC[-5]*100)",
                                NumberFormat = "0.0"
                            },
                            new CellMeta { Row = 4, Column = 7, Text = "\u2264", DisplayText = "\u2264" },
                            new CellMeta
                            {
                                Row = 4,
                                Column = 8,
                                Text = "\u00B12 ppm",
                                DisplayText = "\u00B12 ppm",
                                Formula = "=IF($I$4<=10,\"\u00B12 ppm\",\"\u226410%FS\")",
                                FormulaR1C1 = "=IF(RC[1]<=10,\"\u00B12 ppm\",\"\u226410%FS\")"
                            },
                            new CellMeta
                            {
                                Row = 4,
                                Column = 9,
                                Text = "10 ppm",
                                DisplayText = "10 ppm",
                                RawValueText = "10",
                                Formula = "=K4",
                                FormulaR1C1 = "=RC[2]",
                                NumberFormat = "0 \"ppm\""
                            },
                            new CellMeta { Row = 4, Column = 10, Text = "\u5408\u683C", Formula = "=IF(ABS(F4)<=10,\"\u5408\u683C\",\"\u4E0D\u5408\u683C\")", FormulaR1C1 = "=IF(ABS(RC[-4])<=10,\"\u5408\u683C\",\"\u4E0D\u5408\u683C\")" }
                        }
                    }
                }
            };
        }

        private static TemplateRegionMapping BuildTemplateDefinitionMapping()
        {
            return new TemplateRegionMapping
            {
                ProjectName = "\u793A\u503C\u8BEF\u5DEE",
                SectionRange = new CellRange { SheetName = "Sheet1", StartRow = 1, EndRow = 4, StartColumn = 1, EndColumn = 10 },
                StandardValueRange = RangeAt(4, 1),
                MeasurementValueRange = new CellRange { SheetName = "Sheet1", StartRow = 4, EndRow = 4, StartColumn = 3, EndColumn = 4 },
                AverageValueRange = RangeAt(4, 5),
                ErrorValueRange = RangeAt(4, 6),
                TechnicalRequirementRange = new CellRange { SheetName = "Sheet1", StartRow = 4, EndRow = 4, StartColumn = 7, EndColumn = 8 },
                RangeValueRange = RangeAt(4, 9),
                ResultRange = RangeAt(4, 10)
            };
        }

        private static WorkbookSnapshot BuildSnapshot(double standardValue, double measurementValue)
        {
            return new WorkbookSnapshot
            {
                WorkbookName = "book.xlsx",
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        UsedRangeShape = "A1:D8",
                        Headers = new List<HeaderPath>
                        {
                            new HeaderPath { Column = 3, Levels = new List<string> { "测量值", "1" } }
                        },
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 1, Column = 1, Text = "示值误差" },
                            new CellMeta { Row = 1, Column = 3, Text = "测量值" },
                            new CellMeta { Row = 5, Column = 1, Text = standardValue.ToString() },
                            new CellMeta { Row = 5, Column = 3, Text = measurementValue.ToString() },
                            new CellMeta { Row = 5, Column = 4, Formula = "=C5-A5" },
                            new CellMeta { Row = 6, Column = 1, Text = "±0.5" }
                        }
                    }
                }
            };
        }

        private static CellRange Range(string address)
        {
            return new CellRange
            {
                SheetName = "Sheet1",
                StartRow = 5,
                EndRow = 5,
                StartColumn = address.StartsWith("C", StringComparison.OrdinalIgnoreCase) ? 3 : 4,
                EndColumn = address.StartsWith("C", StringComparison.OrdinalIgnoreCase) ? 3 : 4
            };
        }

        private static CellRange RangeAt(int row, int column)
        {
            return new CellRange
            {
                SheetName = "Sheet1",
                StartRow = row,
                EndRow = row,
                StartColumn = column,
                EndColumn = column
            };
        }

        private static CellRange RangeSpan(int startRow, int endRow, int column)
        {
            return new CellRange
            {
                SheetName = "Sheet1",
                StartRow = startRow,
                EndRow = endRow,
                StartColumn = column,
                EndColumn = column
            };
        }

        private static LocalTemplateRuleCacheRepository CreateRepository()
        {
            var path = Path.Combine(Path.GetTempPath(), "ExcelCalibrationAddin.Tests", Guid.NewGuid().ToString("N") + ".sqlite");
            var repository = new LocalTemplateRuleCacheRepository(path);
            repository.Initialize();
            return repository;
        }

        private static TemplateFingerprint Fingerprint(string exactFingerprint)
        {
            return new TemplateFingerprint
            {
                ExactFingerprint = exactFingerprint,
                FuzzyFingerprint = exactFingerprint + "-fuzzy",
                Title = "Template",
                SheetNames = new List<string> { "Sheet1" },
                HeaderTexts = new List<string> { "Standard", "Measurement" }
            };
        }

        private static MeasurementRule Rule(string name, double standardValue)
        {
            return new MeasurementRule
            {
                FieldName = name,
                TargetRange = Range("C5:C5"),
                FixedStandardValue = standardValue,
                FixedMpe = 0.5,
                ErrorSource = new ParameterSource { Range = Range("D5:D5") },
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                WritableCells = new List<CellAddress> { new CellAddress { Row = 5, Column = 3 } }
            };
        }

        private sealed class RecordingWorkbookWriter : IWorkbookWriter
        {
            public int WriteCount { get; private set; }
            public CellRange LastRange { get; private set; }
            public IReadOnlyList<string> LastValues { get; private set; } = new List<string>();

            public void Write(CellRange range, IReadOnlyList<string> values)
            {
                Write(range, null, values);
            }

            public void Write(CellRange range, IReadOnlyList<CellAddress> writableCells, IReadOnlyList<string> values)
            {
                WriteCount++;
                LastRange = range;
                LastValues = values;
            }
        }

        private sealed class StaticSnapshotProvider : IWorkbookSnapshotProvider
        {
            private readonly WorkbookSnapshot _snapshot;

            public StaticSnapshotProvider(WorkbookSnapshot snapshot)
            {
                _snapshot = snapshot;
            }

            public WorkbookSnapshot Capture()
            {
                return _snapshot;
            }

            public WorkbookSnapshot Capture(IEnumerable<CellRange> ranges)
            {
                return _snapshot;
            }

            public string GetActiveSheetName()
            {
                return _snapshot?.Sheets?.FirstOrDefault()?.Name ?? string.Empty;
            }
        }

        private sealed class StaticJsonHandler : HttpMessageHandler
        {
            private readonly string _response;

            public StaticJsonHandler(string response)
            {
                _response = response;
            }

            public string LastRequestBody { get; private set; } = string.Empty;

            protected override Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request,
                CancellationToken cancellationToken)
            {
                LastRequestBody = request.Content == null
                    ? string.Empty
                    : request.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(_response)
                });
            }
        }

        private sealed class SequenceJsonHandler : HttpMessageHandler
        {
            private readonly Queue<string> _responses;

            public SequenceJsonHandler(params string[] responses)
            {
                _responses = new Queue<string>(responses ?? Array.Empty<string>());
            }

            protected override Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request,
                CancellationToken cancellationToken)
            {
                var response = _responses.Count > 0 ? _responses.Dequeue() : "{}";
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(response)
                });
            }
        }
    }
}
