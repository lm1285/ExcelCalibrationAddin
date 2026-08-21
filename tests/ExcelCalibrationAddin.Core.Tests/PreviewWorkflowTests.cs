using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Services;
using ExcelCalibrationAddin.Host.UseCases;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelCalibrationAddin.Core.Tests
{
    public sealed partial class CoreBehaviorTests
    {
        [TestMethod]
        public void GenerationWithoutErrorFormulaWritesMeasurementsAndReturnsWarning()
        {
            var writer = new RecordingWorkbookWriter();
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(11)),
                new GenerationConfiguration(),
                writer,
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = Range("C5:C5"),
                FixedStandardValue = 10,
                FixedMpe = 0.5,
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                WritableCells = new List<CellAddress> { new CellAddress { Row = 5, Column = 3 } },
                ErrorFormula = new ErrorFormulaInfo { HasFormula = false }
            };

            var result = useCase.WritePreResolved(new[] { rule });

            Assert.AreEqual(1, writer.WriteCount);
            Assert.AreEqual("Sheet1", writer.LastRange.SheetName);
            Assert.AreEqual(3, writer.LastRange.StartColumn);
            Assert.AreEqual(1, writer.LastValues.Count);
            Assert.AreEqual(1, result.WarningMessages.Count);
            StringAssert.Contains(result.WarningMessages[0], "示值误差");
        }

        [TestMethod]
        public void GenerationWithErrorFormulaDoesNotReturnWarning()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(111)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = Rule("带公式示值误差", 10);
            rule.ErrorFormula = new ErrorFormulaInfo
            {
                HasFormula = true,
                Formula = "=C5-A5"
            };

            var result = useCase.WritePreResolved(new[] { rule });

            Assert.AreEqual(0, result.WarningMessages.Count);
        }

        [DataTestMethod]
        [DataRow(180d, 5d)]
        [DataRow(181d, 7d)]
        public void ResponseTimeManualStandardValueOverridesMpeGeneration(double mpe, double maximumSpread)
        {
            var manualStandardValue = mpe + 20d;
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(41)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "响应时间",
                TargetRange = Range("C5:H5"),
                FixedStandardValue = manualStandardValue,
                ManualStandardValues = new List<ManualStandardValue>
                {
                    new ManualStandardValue { PointIndex = 1, Value = manualStandardValue }
                },
                FixedMpe = mpe,
                RequirementOperator = TechnicalRequirementOperator.LessThanOrEqual,
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                WritableCells = Enumerable.Range(3, 6)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.All(value => Math.Abs(value - manualStandardValue) <= maximumSpread / 2d + 1e-12));
            Assert.IsTrue(preview.RawValues.Max() - preview.RawValues.Min() <= maximumSpread + 1e-12);
        }

        [TestMethod]
        public void PlusMinusResponseTimeKeepsOriginalUpperLimitGeneration()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(42)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "响应时间",
                TargetRange = Range("C5:E5"),
                FixedStandardValue = 10,
                ManualStandardValues = new List<ManualStandardValue>
                {
                    new ManualStandardValue { PointIndex = 1, Value = 10 }
                },
                FixedMpe = 180,
                RequirementOperator = TechnicalRequirementOperator.None,
                MpeSource = new ParameterSource { ValuePattern = "mpe:absolute:scale=1:op=plusminus" },
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                WritableCells = Enumerable.Range(3, 3)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.All(value => value >= 36 && value <= 144));
        }

        [TestMethod]
        public void ResponseTimeWithoutStandardStillUsesFixedMaximumSpread()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(43)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "响应时间",
                TargetRange = Range("C5:H5"),
                FixedMpe = 180,
                RequirementOperator = TechnicalRequirementOperator.LessThanOrEqual,
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                WritableCells = Enumerable.Range(3, 6)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.Max() - preview.RawValues.Min() <= 5d + 1e-12);
        }

        [TestMethod]
        public void ResponseTimeManualRangeIsAHardGenerationBoundary()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(44)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "响应时间",
                TargetRange = Range("C5:E5"),
                FixedStandardValue = 25,
                ManualStandardValues = new List<ManualStandardValue>
                {
                    new ManualStandardValue { PointIndex = 1, Value = 25 }
                },
                MeasurementLowerBound = 20,
                MeasurementUpperBound = 30,
                FixedMpe = 60,
                RequirementOperator = TechnicalRequirementOperator.LessThanOrEqual,
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                WritableCells = Enumerable.Range(3, 3)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.All(value => value >= 20 && value <= 30));
            Assert.IsTrue(preview.RawValues.Max() - preview.RawValues.Min() <= 5d + 1e-12);
        }

        [TestMethod]
        public void TwoManualStandardsDefineRangeForSingleResponseTimeRow()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(45)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "响应时间",
                TargetRange = Range("C5:E5"),
                FixedStandardValue = 20,
                ManualStandardValues = new List<ManualStandardValue>
                {
                    new ManualStandardValue { PointIndex = 1, Value = 20 },
                    new ManualStandardValue { PointIndex = 2, Value = 30 }
                },
                FixedMpe = 60,
                RequirementOperator = TechnicalRequirementOperator.LessThanOrEqual,
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                WritableCells = Enumerable.Range(3, 3)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.All(value => value >= 20 && value <= 30));
            Assert.IsTrue(preview.RawValues.Max() - preview.RawValues.Min() <= 5d + 1e-12);
        }

        [DataTestMethod]
        [DataRow("20-30", 20d, 30d)]
        [DataRow("30~20", 20d, 30d)]
        [DataRow("20～30", 20d, 30d)]
        [DataRow("20至30", 20d, 30d)]
        public void ManualStandardRangeParserAcceptsCommonRangeFormats(
            string text,
            double expectedLowerBound,
            double expectedUpperBound)
        {
            Assert.IsTrue(ManualStandardValueRangeParser.TryParse(text, out var lowerBound, out var upperBound));
            Assert.AreEqual(expectedLowerBound, lowerBound, 1e-12);
            Assert.AreEqual(expectedUpperBound, upperBound, 1e-12);
        }

        [TestMethod]
        public void ParameterResolverKeepsManualStandardValueInsteadOfWorksheetValue()
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
                            new CellMeta { Row = 5, Column = 1, Text = "25" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "响应时间",
                StandardValueSource = new ParameterSource { Range = Range("A5:A5") },
                FixedStandardValue = 100,
                ManualStandardValues = new List<ManualStandardValue>
                {
                    new ManualStandardValue { PointIndex = 1, Value = 100 }
                }
            };

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(100d, resolved.FixedStandardValue.Value, 1e-12);
            Assert.AreEqual(100d, resolved.ManualStandardValues.Single().Value.Value, 1e-12);
        }

        [TestMethod]
        public void GenerationWriteResultRemovesDuplicateAndBlankWarnings()
        {
            var result = GenerationWriteResult.FromPreviews(new[]
            {
                new RulePreview { WarningMessages = new[] { "缺少误差公式", "" } },
                new RulePreview { WarningMessages = new[] { "缺少误差公式", "需要人工确认" } },
                null
            });

            CollectionAssert.AreEqual(
                new[] { "缺少误差公式", "需要人工确认" },
                new List<string>(result.WarningMessages));
        }

        [TestMethod]
        public void StandardMeasurementsRemainVisiblyDifferentWhenResolutionOptionIsDisabled()
        {
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = Enumerable.Range(3, 3)
                            .Select(column => new CellMeta { Row = 5, Column = column, NumberFormat = "0" })
                            .ToList()
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = new CellRange
                {
                    SheetName = "Sheet1",
                    StartRow = 5,
                    EndRow = 5,
                    StartColumn = 3,
                    EndColumn = 5
                },
                FixedStandardValue = 10,
                FixedMpe = 5,
                PositiveDirectionOnly = true,
                FormatRule = new FormatRule { DecimalPlaces = 0 }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(91)),
                new GenerationConfiguration { UseDecimalPlacesForResolution = false },
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.DisplayValues.Distinct().Count() > 1);
        }

        [TestMethod]
        public void RepeatabilityUsesErrorColumnPrecisionToAvoidZeroDisplay()
        {
            var cells = Enumerable.Range(3, 6)
                .Select(column => new CellMeta { Row = 5, Column = column, NumberFormat = "0.00" })
                .ToList();
            cells.Add(new CellMeta { Row = 5, Column = 9, NumberFormat = "0" });
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot { Name = "Sheet1", Cells = cells }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "重复性",
                TargetRange = new CellRange
                {
                    SheetName = "Sheet1",
                    StartRow = 5,
                    EndRow = 5,
                    StartColumn = 3,
                    EndColumn = 8
                },
                ErrorSource = new ParameterSource { Range = RangeAt(5, 9) },
                FixedStandardValue = 40,
                FixedMpe = 2,
                FormatRule = new FormatRule { DecimalPlaces = 2 }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(92)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            var average = preview.RawValues.Average();
            var standardDeviation = Math.Sqrt(
                preview.RawValues.Sum(value => Math.Pow(value - average, 2)) /
                (preview.RawValues.Count - 1));

            Assert.IsTrue(preview.RawValues.Distinct().Count() > 1);
            Assert.IsTrue(Math.Round(standardDeviation / Math.Abs(average) * 100d, 0) > 0);
        }

        [TestMethod]
        public void RepeatabilityWithIntegerMeasurementsDoesNotCollapseToZero()
        {
            var cells = Enumerable.Range(3, 6)
                .Select(column => new CellMeta { Row = 5, Column = column, NumberFormat = "0" })
                .ToList();
            cells.Add(new CellMeta { Row = 5, Column = 2, Text = "40" });
            cells.Add(new CellMeta { Row = 5, Column = 9, NumberFormat = "0.0" });
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot { Name = "Sheet1", Cells = cells }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "重复性",
                TargetRange = new CellRange
                {
                    SheetName = "Sheet1",
                    StartRow = 5,
                    EndRow = 5,
                    StartColumn = 3,
                    EndColumn = 8
                },
                ErrorSource = new ParameterSource { Range = RangeAt(5, 9) },
                FixedStandardValue = 40,
                FixedMpe = 2,
                FormatRule = new FormatRule { DecimalPlaces = 0 }
            };
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(7)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(snapshot),
                null);

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.Distinct().Count() > 1);
            var average = preview.RawValues.Average();
            var standardDeviation = Math.Sqrt(
                preview.RawValues.Sum(value => Math.Pow(value - average, 2)) /
                (preview.RawValues.Count - 1));
            Assert.IsTrue(standardDeviation > 0);
        }
    }
}
