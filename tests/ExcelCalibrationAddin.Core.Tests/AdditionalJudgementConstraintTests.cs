using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Recognition;
using ExcelCalibrationAddin.Host.Services;
using ExcelCalibrationAddin.Host.UseCases;
using ExcelCalibrationAddin.Host.Generation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelCalibrationAddin.Core.Tests
{
    [TestClass]
    public sealed class AdditionalJudgementConstraintTests
    {
        [TestMethod]
        public void DraftBuilderKeepsSecondaryJudgementColumnsOnTheSameCalibrationItem()
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
                            new CellMeta { Row = 2, Column = 1, Text = "标准值" },
                            new CellMeta { Row = 2, Column = 2, Text = "测量值" },
                            new CellMeta { Row = 2, Column = 3, Text = "示值误差" },
                            new CellMeta { Row = 2, Column = 4, Text = "技术要求" },
                            new CellMeta { Row = 2, Column = 5, Text = "结论" },
                            new CellMeta { Row = 2, Column = 6, Text = "相对误差" },
                            new CellMeta { Row = 2, Column = 7, Text = "技术要求" },
                            new CellMeta { Row = 2, Column = 8, Text = "结论" },
                            new CellMeta { Row = 3, Column = 1, Text = "100.0", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 2, Text = "100.2", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 3, Text = "0.2", Formula = "=B3-A3", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 4, Text = "±1.0" },
                            new CellMeta { Row = 3, Column = 5, Text = "合格", Formula = "=IF(ABS(C3)<=1,\"合格\",\"不合格\")" },
                            new CellMeta { Row = 3, Column = 6, Text = "0.2%", Formula = "=(B3-A3)/A3", NumberFormat = "0.00%" },
                            new CellMeta { Row = 3, Column = 7, Text = "±1%" },
                            new CellMeta { Row = 3, Column = 8, Text = "合格", Formula = "=IF(ABS(F3)<=0.01,\"合格\",\"不合格\")" }
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
                            EndColumn = 8
                        }
                    }
                }
            };

            var mappings = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter()).BuildMappings(recognition);


            Assert.AreEqual(1, mappings.Count);
            Assert.AreEqual(1, mappings[0].AdditionalJudgementConstraints.Count);
            Assert.AreEqual(6, mappings[0].AdditionalJudgementConstraints[0].ErrorSource.Range.StartColumn);
            Assert.AreEqual(7, mappings[0].AdditionalJudgementConstraints[0].MpeSource.Range.StartColumn);
            Assert.AreEqual(8, mappings[0].AdditionalJudgementConstraints[0].ResultSource.Range.StartColumn);

            var rule = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter())
                .BuildDraftRules(recognition, mappings)
                .Single();
            Assert.AreEqual(1, rule.AdditionalJudgementConstraints.Count);
            Assert.AreEqual(3, rule.ErrorSource.Range.StartColumn);
            Assert.AreEqual(6, rule.AdditionalJudgementConstraints[0].ErrorSource.Range.StartColumn);
        }

        [TestMethod]
        public void ParameterResolverAndStructureAnalyzerPopulateSecondaryConstraintFormulas()
        {
            var snapshot = BuildSecondaryConstraintSnapshot();
            var rule = BuildSecondaryConstraintRule();

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();
            new MeasurementRuleStructureAnalyzer().Apply(snapshot, new[] { resolved });

            Assert.AreEqual(1d, resolved.FixedMpe.GetValueOrDefault(), 1e-12);
            Assert.AreEqual(0.01d, resolved.AdditionalJudgementConstraints[0].FixedMpe.GetValueOrDefault(), 1e-12);
            Assert.IsTrue(resolved.ErrorFormula.HasFormula);
            Assert.IsTrue(resolved.AdditionalJudgementConstraints[0].ErrorFormula.HasFormula);
            StringAssert.Contains(resolved.AdditionalJudgementConstraints[0].ErrorFormula.Formula, "B3-A3");
        }

        [TestMethod]
        public void SecondaryConstraintKeepsTighterTechnicalRequirementAfterResolution()
        {
            var snapshot = BuildSecondaryConstraintSnapshot();
            var rule = BuildSecondaryConstraintRule();

            var resolved = new MeasurementRuleParameterResolver().Apply(snapshot, new[] { rule }).Single();

            Assert.AreEqual(1d, resolved.FixedMpe.GetValueOrDefault(), 1e-12);
            Assert.IsTrue(resolved.AdditionalJudgementConstraints[0].FixedMpe.GetValueOrDefault() < resolved.FixedMpe.GetValueOrDefault());
            Assert.AreEqual(TechnicalRequirementOperator.PlusMinus, resolved.AdditionalJudgementConstraints[0].RequirementOperator);
        }

        [TestMethod]
        public void DraftBuilderPairsThreeErrorAndMpeColumnsBySharedRows()
        {
            var snapshot = BuildMultiConstraintSnapshot(threeMpe: true);
            var recognition = BuildRecognition(snapshot, "示值误差与重复性");

            var mapping = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter())
                .BuildMappings(recognition)
                .Single();


            Assert.IsTrue(string.IsNullOrWhiteSpace(mapping.RecognitionError));
            Assert.AreEqual(2, mapping.AdditionalJudgementConstraints.Count);
            Assert.AreEqual(6, mapping.AdditionalJudgementConstraints[0].ErrorSource.Range.StartColumn);
            Assert.AreEqual(7, mapping.AdditionalJudgementConstraints[0].MpeSource.Range.StartColumn);
            Assert.AreEqual(9, mapping.AdditionalJudgementConstraints[1].ErrorSource.Range.StartColumn);
            Assert.AreEqual(10, mapping.AdditionalJudgementConstraints[1].MpeSource.Range.StartColumn);
        }

        [TestMethod]
        public void DraftBuilderFailsWhenErrorAndMpeCountsDiffer()
        {
            var snapshot = BuildMultiConstraintSnapshot(threeMpe: false);
            var mapping = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter())
                .BuildMappings(BuildRecognition(snapshot, "示值误差与重复性"))
                .Single();


            StringAssert.Contains(mapping.RecognitionError, "数量不一致");
            var rule = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter())
                .BuildDraftRules(BuildRecognition(snapshot, "示值误差与重复性"), new[] { mapping })
                .Single();
            Assert.ThrowsException<System.InvalidOperationException>(() =>
                GenerationRuleValidator.ValidateFormulaDependencies(rule));
        }

        [TestMethod]
        public void DraftBuilderPairsVerticallyStackedErrorAndMpeRows()
        {
            var cells = new List<CellMeta>
            {
                new CellMeta { Row = 1, Column = 1, Text = "示值误差与重复性" },
                new CellMeta { Row = 2, Column = 1, Text = "标准值" },
                new CellMeta { Row = 2, Column = 2, Text = "测量值" },
                new CellMeta { Row = 2, Column = 3, Text = "示值误差" },
                new CellMeta { Row = 2, Column = 4, Text = "技术要求" },
                new CellMeta { Row = 2, Column = 5, Text = "结论" },
                new CellMeta { Row = 3, Column = 1, Text = "100" },
                new CellMeta { Row = 3, Column = 2, Text = "100.2" },
                new CellMeta { Row = 3, Column = 3, Text = "0.2", Formula = "=B3-A3" },
                new CellMeta { Row = 3, Column = 4, Text = "±1" },
                new CellMeta { Row = 3, Column = 5, Text = "合格" },
                new CellMeta { Row = 4, Column = 3, Text = "重复性" },
                new CellMeta { Row = 4, Column = 4, Text = "技术要求" },
                new CellMeta { Row = 4, Column = 5, Text = "结论" },
                new CellMeta { Row = 5, Column = 1, Text = "100" },
                new CellMeta { Row = 5, Column = 2, Text = "100.1" },
                new CellMeta { Row = 5, Column = 3, Text = "0.1", Formula = "=B5-A5" },
                new CellMeta { Row = 5, Column = 4, Text = "±0.5" },
                new CellMeta { Row = 5, Column = 5, Text = "合格" }
            };
            var snapshot = new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot> { new SheetSnapshot { Name = "Sheet1", Cells = cells } }
            };
            var mapping = new MeasurementRuleDraftBuilder(new NumberFormatInterpreter())
                .BuildMappings(BuildRecognition(snapshot, "示值误差与重复性"))
                .Single();

            Assert.IsTrue(string.IsNullOrWhiteSpace(mapping.RecognitionError));
            Assert.AreEqual(1, mapping.AdditionalJudgementConstraints.Count);
            Assert.AreEqual(5, mapping.AdditionalJudgementConstraints[0].ErrorSource.Range.StartRow);
            Assert.AreEqual(5, mapping.AdditionalJudgementConstraints[0].MpeSource.Range.StartRow);
        }

        private static RecognitionResult BuildRecognition(WorkbookSnapshot snapshot, string alias)
        {
            return new RecognitionResult
            {
                Snapshot = snapshot,
                RecognizedFields = new List<RecognizedField>
                {
                    new RecognizedField
                    {
                        Alias = alias,
                        Score = 96,
                        Range = new CellRange { SheetName = "Sheet1", StartRow = 1, EndRow = snapshot.Sheets.First().Cells.Max(cell => cell.Row), StartColumn = 1, EndColumn = 11 }
                    }
                }
            };
        }

        private static WorkbookSnapshot BuildMultiConstraintSnapshot(bool threeMpe)
        {
            var cells = new List<CellMeta>
            {
                new CellMeta { Row = 1, Column = 1, Text = "示值误差与重复性" },
                new CellMeta { Row = 2, Column = 1, Text = "标准值" },
                new CellMeta { Row = 2, Column = 2, Text = "测量值" },
                new CellMeta { Row = 2, Column = 3, Text = "示值误差" },
                new CellMeta { Row = 2, Column = 4, Text = "技术要求" },
                new CellMeta { Row = 2, Column = 5, Text = "结论" },
                new CellMeta { Row = 2, Column = 6, Text = "重复性" },
                new CellMeta { Row = 2, Column = 7, Text = "技术要求" },
                new CellMeta { Row = 2, Column = 8, Text = "结论" },
                new CellMeta { Row = 2, Column = 9, Text = "相对误差" },
                new CellMeta { Row = 2, Column = 10, Text = threeMpe ? "技术要求" : "说明" },
                new CellMeta { Row = 2, Column = 11, Text = "结论" },
                new CellMeta { Row = 3, Column = 1, Text = "100" },
                new CellMeta { Row = 3, Column = 2, Text = "100.2" },
                new CellMeta { Row = 3, Column = 3, Text = "0.2", Formula = "=B3-A3" },
                new CellMeta { Row = 3, Column = 4, Text = "±1" },
                new CellMeta { Row = 3, Column = 5, Text = "合格", Formula = "=IF(ABS(C3)<=D3,\"合格\",\"不合格\")" },
                new CellMeta { Row = 3, Column = 6, Text = "0.1", Formula = "=MAX(B3)-MIN(B3)" },
                new CellMeta { Row = 3, Column = 7, Text = "±0.5" },
                new CellMeta { Row = 3, Column = 8, Text = "合格", Formula = "=IF(ABS(F3)<=G3,\"合格\",\"不合格\")" },
                new CellMeta { Row = 3, Column = 9, Text = "0.3", Formula = "=C3" },
                new CellMeta { Row = 3, Column = 11, Text = "合格", Formula = "=IF(ABS(I3)<=J3,\"合格\",\"不合格\")" }
            };
            if (threeMpe)
            {
                cells.Add(new CellMeta { Row = 3, Column = 10, Text = "±0.8" });
            }

            return new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot> { new SheetSnapshot { Name = "Sheet1", Cells = cells } }
            };
        }

        private static WorkbookSnapshot BuildSecondaryConstraintSnapshot()
        {
            return new WorkbookSnapshot
            {
                Sheets = new List<SheetSnapshot>
                {
                    new SheetSnapshot
                    {
                        Name = "Sheet1",
                        Cells = new List<CellMeta>
                        {
                            new CellMeta { Row = 2, Column = 3, Text = "示值误差" },
                            new CellMeta { Row = 2, Column = 4, Text = "技术要求" },
                            new CellMeta { Row = 2, Column = 6, Text = "相对误差" },
                            new CellMeta { Row = 2, Column = 7, Text = "技术要求" },
                            new CellMeta { Row = 3, Column = 1, Text = "100", RawValueText = "100" },
                            new CellMeta { Row = 3, Column = 2, Text = "100.2", RawValueText = "100.2", NumberFormat = "0.0" },
                            new CellMeta { Row = 3, Column = 3, Text = "0.2", Formula = "=B3-A3", RawValueText = "0.2" },
                            new CellMeta { Row = 3, Column = 4, Text = "±1.0" },
                            new CellMeta { Row = 3, Column = 6, Text = "0.002", Formula = "=(B3-A3)/A3", RawValueText = "0.002" },
                            new CellMeta { Row = 3, Column = 7, Text = "±0.01" }
                        }
                    }
                }
            };
        }

        private static MeasurementRule BuildSecondaryConstraintRule()
        {
            return new MeasurementRule
            {
                FieldName = "示值误差",
                TargetRange = RangeAt(3, 2),
                StandardValueSource = new ParameterSource { Range = RangeAt(3, 1) },
                ErrorSource = new ParameterSource { Name = "误差", Range = RangeAt(3, 3) },
                MpeSource = new ParameterSource { Name = "技术要求", Range = RangeAt(3, 4) },
                AdditionalJudgementConstraints = new List<MeasurementJudgementConstraint>
                {
                    new MeasurementJudgementConstraint
                    {
                        Name = "附属判定1",
                        ErrorSource = new ParameterSource { Name = "误差", Range = RangeAt(3, 6) },
                        MpeSource = new ParameterSource { Name = "技术要求", Range = RangeAt(3, 7) }
                    }
                }
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
    }
}
