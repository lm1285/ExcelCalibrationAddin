using System;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelCalibrationAddin.Core.Tests
{
    [TestClass]
    public sealed class MeasurementSeriesGeneratorTests
    {
        [TestMethod]
        public void RepeatabilityValuesStayAroundCenterAndWithinConfiguredSpread()
        {
            var configuration = new GenerationConfiguration
            {
                AbsoluteErrorMinimumCoefficient = 0.2,
                AbsoluteErrorMaximumCoefficient = 0.8
            };
            var generator = new MeasurementSeriesGenerator(new Random(17));

            var values = generator.GenerateRepeatabilityValues(
                new MeasurementRule(),
                standardValue: 100,
                centerValue: 104,
                toleranceRatio: 0.02,
                valueCount: 6,
                decimalPlaces: 2,
                configuration: configuration);

            Assert.AreEqual(6, values.Count);
            Assert.IsTrue(values.All(value => value >= 103.2 && value <= 104.8));
            Assert.IsTrue(values.Max() - values.Min() <= 1.6 + 1e-12);
            Assert.IsTrue(values.Distinct().Count() > 1);
        }

        [TestMethod]
        public void RepeatabilityWithIntegerResolutionCannotCollapseToZero()
        {
            var generator = new MeasurementSeriesGenerator(new Random(19));

            var values = generator.GenerateRepeatabilityValues(
                new MeasurementRule(),
                standardValue: 40,
                centerValue: 38,
                toleranceRatio: 0.02,
                valueCount: 6,
                decimalPlaces: 0,
                configuration: new GenerationConfiguration());

            Assert.AreEqual(6, values.Count);
            Assert.IsTrue(values.Distinct().Count() > 1);
            Assert.IsTrue(values.Max() - values.Min() > 0);
        }

        [TestMethod]
        public void UpperLimitValuesHonorRuleCoefficientOverrideAndResolution()
        {
            var rule = new MeasurementRule
            {
                GenerationCoefficientOverride = new MeasurementGenerationCoefficientOverride
                {
                    AbsoluteMinimumCoefficient = 0.4,
                    AbsoluteMaximumCoefficient = 0.6
                }
            };
            var configuration = new GenerationConfiguration { DefaultDistribution = "Uniform" };
            var generator = new MeasurementSeriesGenerator(new Random(23));

            var values = generator.GenerateUpperLimitValues(
                rule,
                upperLimit: 10,
                valueCount: 20,
                decimalPlaces: 2,
                configuration: configuration);

            Assert.AreEqual(20, values.Count);
            Assert.IsTrue(values.All(value => value >= 4 && value <= 6));
            Assert.IsTrue(values.All(value => Math.Abs(value - Math.Round(value, 2)) <= 1e-12));
        }

        [DataTestMethod]
        [DataRow(TechnicalRequirementOperator.LessThan, 25d, 35d)]
        [DataRow(TechnicalRequirementOperator.LessThanOrEqual, 25d, 35d)]
        [DataRow(TechnicalRequirementOperator.GreaterThan, 120d, 140d)]
        [DataRow(TechnicalRequirementOperator.GreaterThanOrEqual, 120d, 140d)]
        public void ResponseTimeValuesFollowRequirementOperator(
            TechnicalRequirementOperator requirementOperator,
            double expectedMinimum,
            double expectedMaximum)
        {
            var generator = new MeasurementSeriesGenerator(new Random(31));

            var values = generator.GenerateResponseTimeValues(
                requirement: 100d,
                requirementOperator: requirementOperator,
                valueCount: 50,
                decimalPlaces: 2,
                configuration: new GenerationConfiguration { DefaultDistribution = "Uniform" });

            Assert.AreEqual(50, values.Count);
            Assert.IsTrue(values.All(value => value >= expectedMinimum && value <= expectedMaximum));
        }

        [TestMethod]
        public void ResponseTimeValuesPreserveDecimalPlaces()
        {
            var generator = new MeasurementSeriesGenerator(new Random(32));

            var values = generator.GenerateResponseTimeValues(
                requirement: 60,
                requirementOperator: TechnicalRequirementOperator.LessThanOrEqual,
                valueCount: 50,
                decimalPlaces: 2,
                configuration: new GenerationConfiguration { DefaultDistribution = "Uniform" });

            Assert.IsTrue(values.All(value => Math.Abs(value - Math.Round(value, 2)) <= 1e-12));
        }
    }
}
