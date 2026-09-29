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

        [TestMethod]
        public void AdditionalMaxMinusMinConstraintLimitsGeneratedMeasurementSpread()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(411)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "波长示值误差与重复性的校准",
                TargetRange = Range("G6:O6"),
                FixedStandardValue = 241.3,
                FixedMpe = 2,
                ErrorSource = new ParameterSource { Range = Range("S6:U6") },
                ErrorFormula = new ErrorFormulaInfo { HasFormula = true, Formula = "=P6-D6" },
                FormatRule = new FormatRule { DecimalPlaces = 1 },
                WritableCells = new[] { 7, 10, 13 }
                    .Select(column => new CellAddress { Row = 6, Column = column })
                    .ToList(),
                AdditionalJudgementConstraints = new List<MeasurementJudgementConstraint>
                {
                    new MeasurementJudgementConstraint
                    {
                        Name = "重复性",
                        ErrorSource = new ParameterSource { Range = Range("Y6:AA6") },
                        FixedMpe = 1,
                        ErrorType = ErrorType.Absolute,
                        ErrorFormula = new ErrorFormulaInfo
                        {
                            HasFormula = true,
                            Formula = "=MAX(G6:O6)-MIN(G6:O6)",
                            Scale = ErrorFormulaScale.Absolute
                        }
                    }
                }
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.AreEqual(3, preview.RawValues.Count);
            Assert.IsTrue(preview.RawValues.Max() - preview.RawValues.Min() <= 1.0000000001,
                "Generated measurements must satisfy the MAX-MIN repeatability bound. Values=" +
                string.Join(",", preview.RawValues));
        }

        [TestMethod]
        public void GenerationRejectsUnresolvedFormulaDependencies()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(112)),
                new GenerationConfiguration(),
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = Rule("结构化引用项目", 10);
            rule.ErrorSource = new ParameterSource { Range = Range("D5:D5") };
            rule.ErrorFormula = new ErrorFormulaInfo
            {
                HasFormula = true,
                Formula = "=Table1[Value]-C5",
                UnresolvedDependencies = new List<string> { "结构化引用: Table1[Value]" }
            };

            var exception = Assert.ThrowsException<System.InvalidOperationException>(() =>
                useCase.PreviewPreResolved(new[] { rule }));

            StringAssert.Contains(exception.Message, "无法安全解析的公式依赖");
            StringAssert.Contains(exception.Message, "Table1[Value]");
        }

        [TestMethod]
        public void SelectedSampleDataProvidesCandidatesBeforeFormulaValidation()
        {
            var configuration = new GenerationConfiguration
            {
                PositiveErrorMinimumCoefficient = 0.7,
                PositiveErrorMaximumCoefficient = 0.8,
                NegativeErrorMinimumCoefficient = 0.7,
                NegativeErrorMaximumCoefficient = 0.8,
                AbsoluteErrorMinimumCoefficient = 0.7,
                AbsoluteErrorMaximumCoefficient = 0.8
            };
            var useCase = new GenerateMeasurementUseCase(
                value => new MeasurementValueGenerator(value, new Random(113)),
                configuration,
                new RecordingWorkbookWriter(),
                new StaticSnapshotProvider(new WorkbookSnapshot
                {
                    Sheets = new List<SheetSnapshot>
                    {
                        new SheetSnapshot
                        {
                            Name = "Sheet1",
                            Cells = Enumerable.Range(3, 3)
                                .Select(column => new CellMeta { Row = 5, Column = column, NumberFormat = "0.00" })
                                .ToList()
                        }
                    }
                }),
                null);
            useCase.SetSampleDataPoints(new[]
            {
                new SampleDataPoint
                {
                    CalibrationItemName = "样本项目",
                    StandardValue = 10,
                    DecimalPlaces = 2,
                    MeasurementValues = new List<double> { 10.36, 10.37, 10.38 }
                }
            });
            var rule = Rule("样本项目", 10);
            rule.TargetRange = Range("C5:E5");
            rule.WritableCells = Enumerable.Range(3, 3)
                .Select(column => new CellAddress { Row = 5, Column = column })
                .ToList();

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.All(value => value >= 10.35 && value <= 10.39));
        }

        [DataTestMethod]
        [DataRow(TechnicalRequirementOperator.LessThan, 45d, 63d)]
        [DataRow(TechnicalRequirementOperator.LessThanOrEqual, 45d, 63d)]
        [DataRow(TechnicalRequirementOperator.GreaterThan, 216d, 252d)]
        [DataRow(TechnicalRequirementOperator.GreaterThanOrEqual, 216d, 252d)]
        public void SecondsRequirementUsesConfirmedRangeRegardlessOfProjectName(
            TechnicalRequirementOperator requirementOperator,
            double expectedMinimum,
            double expectedMaximum)
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(41)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "动作时间",
                TargetRange = Range("C5:H5"),
                FixedMpe = 180d,
                RequirementOperator = requirementOperator,
                MpeSource = new ParameterSource { ValuePattern = "mpe:absolute:scale=1:unit=s" },
                FormatRule = new FormatRule { DecimalPlaces = 2, UnitSuffix = "s" },
                WritableCells = Enumerable.Range(3, 6)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.All(value => value >= expectedMinimum && value <= expectedMaximum));
        }

        [TestMethod]
        public void TimeKeywordInfersLessOrEqualFromResultFormula()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(46)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "2.4、响应时间",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 20, EndRow = 20, StartColumn = 5, EndColumn = 16 },
                FixedMpe = 60,
                RequirementOperator = TechnicalRequirementOperator.None,
                AverageSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 20, EndRow = 20, StartColumn = 17, EndColumn = 17 } },
                ErrorSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 20, EndRow = 20, StartColumn = 17, EndColumn = 17 } },
                MpeSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 20, EndRow = 20, StartColumn = 23, EndColumn = 23 } },
                ResultSource = new ParameterSource { Range = new CellRange { SheetName = "Sheet1", StartRow = 20, EndRow = 20, StartColumn = 30, EndColumn = 30 } },
                FormatRule = new FormatRule { DecimalPlaces = 2 },
                WritableCells = new List<CellAddress>
                {
                    new CellAddress { Row = 20, Column = 5 },
                    new CellAddress { Row = 20, Column = 9 },
                    new CellAddress { Row = 20, Column = 13 }
                },
                ErrorFormula = new ErrorFormulaInfo
                {
                    HasFormula = true,
                    Formula = "=AVERAGE(E20:P20)",
                    ResultFormula = "=IF(W20=\"/\",\"P\",IF(Q20<=W20,\"P\",\"F\"))",
                    ResultFormulaResolved = true
                },
                TemplateDefinition = new TemplateFieldDefinition
                {
                    ProjectName = "2.4、响应时间",
                    Regions = new List<TemplateRegionDefinition>
                    {
                        new TemplateRegionDefinition
                        {
                            Role = TemplateRegionRole.TechnicalRequirement,
                            Unit = "s",
                            Units = new List<string> { "s" }
                        }
                    }
                }
            };

            Assert.IsTrue(ExcelCalibrationAddin.Host.Generation.GenerationRuleValidator.IsUpperLimitRule(rule));
            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            Assert.IsTrue(preview.RawValues.All(value => value >= 15 && value <= 21));
        }

        [TestMethod]
        public void TimeKeywordGreaterThanUsesHighRequirementRange()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(47)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "响应时间",
                TargetRange = Range("C5:H5"),
                FixedMpe = 100,
                RequirementOperator = TechnicalRequirementOperator.GreaterThan,
                MpeSource = new ParameterSource { ValuePattern = "mpe:absolute:scale=1:op=greaterthan:unit=s" },
                FormatRule = new FormatRule { DecimalPlaces = 2, UnitSuffix = "s" },
                WritableCells = Enumerable.Range(3, 6)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            Assert.IsTrue(preview.RawValues.All(value => value >= 120 && value <= 140));
        }

        [TestMethod]
        public void SecondsRequirementWithoutTimeKeywordDoesNotUseUpperLimitPath()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(48)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = Range("C5:C5"),
                FixedStandardValue = 10,
                FixedMpe = 180,
                RequirementOperator = TechnicalRequirementOperator.LessThanOrEqual,
                MpeSource = new ParameterSource { ValuePattern = "mpe:absolute:scale=1:unit=s" },
                FormatRule = new FormatRule { DecimalPlaces = 2, UnitSuffix = "s" },
                WritableCells = new List<CellAddress> { new CellAddress { Row = 5, Column = 3 } }
            };

            Assert.IsFalse(ExcelCalibrationAddin.Host.Generation.GenerationRuleValidator.IsUpperLimitRule(rule));
            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();
            Assert.AreEqual(1, preview.RawValues.Count);
            Assert.IsTrue(preview.RawValues[0] >= 10 && preview.RawValues[0] <= 10 + 180);
        }

        [TestMethod]
        public void PlusMinusResponseTimeDoesNotUseSpecialResponseTimePath()
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
                ErrorSource = new ParameterSource { Range = Range("D5:D5") },
                MpeSource = new ParameterSource { ValuePattern = "mpe:absolute:scale=1:op=plusminus:unit=s" },
                FormatRule = new FormatRule { DecimalPlaces = 2, UnitSuffix = "s" },
                WritableCells = Enumerable.Range(3, 3)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsFalse(ExcelCalibrationAddin.Host.Generation.GenerationRuleValidator.IsUpperLimitRule(rule));
            Assert.AreEqual(3, preview.RawValues.Count);
        }

        [TestMethod]
        public void ResponseTimeWithoutSecondsUnitDoesNotUseSpecialResponseTimePath()
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
                FixedStandardValue = 100,
                ErrorSource = new ParameterSource { Range = Range("D5:D5") },
                RequirementOperator = TechnicalRequirementOperator.LessThanOrEqual,
                FormatRule = new FormatRule { DecimalPlaces = 2, UnitSuffix = "ms" },
                WritableCells = Enumerable.Range(3, 6)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsFalse(ExcelCalibrationAddin.Host.Generation.GenerationRuleValidator.IsUpperLimitRule(rule));
            Assert.AreEqual(6, preview.RawValues.Count);
        }

        [TestMethod]
        public void ResponseTimeManualRangeDoesNotOverrideRequirementRule()
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
                MpeSource = new ParameterSource { ValuePattern = "mpe:absolute:scale=1:unit=s" },
                FormatRule = new FormatRule { DecimalPlaces = 2, UnitSuffix = "s" },
                WritableCells = Enumerable.Range(3, 3)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.All(value => value >= 15 && value <= 21));
        }

        [TestMethod]
        public void MultipleManualStandardsDoNotOverrideResponseTimeRequirementRule()
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
                MpeSource = new ParameterSource { ValuePattern = "mpe:absolute:scale=1:unit=s" },
                FormatRule = new FormatRule { DecimalPlaces = 2, UnitSuffix = "s" },
                WritableCells = Enumerable.Range(3, 3)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.All(value => value >= 15 && value <= 21));
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
        public void NearbyPercentLelHeaderDoesNotScaleSecondsRequirement()
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
                            new CellMeta { Row = 4, Column = 1, Text = "标准值\n(%LEL)" },
                            new CellMeta { Row = 4, Column = 3, Text = "测量值/s" },
                            new CellMeta { Row = 4, Column = 6, Text = "技术要求" },
                            new CellMeta { Row = 5, Column = 1, Text = "40.0" },
                            new CellMeta { Row = 5, Column = 6, Text = "扩散式：≤60s", DisplayText = "扩散式：≤60s" }
                        }
                    }
                }
            };
            var rule = new MeasurementRule
            {
                FieldName = "七、响应时间",
                TargetRange = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 3, EndColumn = 5 },
                MpeSource = new ParameterSource
                {
                    Range = new CellRange { SheetName = "Sheet1", StartRow = 5, EndRow = 5, StartColumn = 6, EndColumn = 6 }
                },
                FormatRule = new FormatRule { DecimalPlaces = 2, UnitSuffix = "s" },
                WritableCells = Enumerable.Range(3, 3)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(60d, resolved.FixedMpe.GetValueOrDefault(), 1e-12);
            Assert.AreEqual(TechnicalRequirementOperator.LessThanOrEqual, resolved.RequirementOperator);
            Assert.IsTrue(ExcelCalibrationAddin.Host.Generation.GenerationRuleValidator.IsUpperLimitRule(resolved));
        }

        [TestMethod]
        public void SecondsUpperLimitRestoresPercentScaledRequirement()
        {
            var useCase = new GenerateMeasurementUseCase(
                configuration => new MeasurementValueGenerator(configuration, new Random(49)),
                new GenerationConfiguration { DefaultDistribution = "Uniform" },
                new RecordingWorkbookWriter(),
                null,
                null);
            var rule = new MeasurementRule
            {
                FieldName = "响应时间",
                TargetRange = Range("C5:E5"),
                FixedMpe = 0.6,
                RequirementOperator = TechnicalRequirementOperator.LessThanOrEqual,
                MpeSource = new ParameterSource { ValuePattern = "mpe:absolute:scale=0.01:op=lessthanorequal:unit=s" },
                FormatRule = new FormatRule { DecimalPlaces = 2, UnitSuffix = "s" },
                WritableCells = Enumerable.Range(3, 3)
                    .Select(column => new CellAddress { Row = 5, Column = column })
                    .ToList()
            };

            var preview = useCase.PreviewPreResolved(new[] { rule }).Single();

            Assert.IsTrue(preview.RawValues.All(value => value >= 15 && value <= 21));
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
