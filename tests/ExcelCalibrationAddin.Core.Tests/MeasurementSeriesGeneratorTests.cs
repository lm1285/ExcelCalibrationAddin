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
        [DataRow(180d, 5d)]
        [DataRow(180.01d, 7d)]
        public void ResponseTimeValuesUseFixedMaximumSpread(double mpe, double expectedMaximumSpread)
        {
            var generator = new MeasurementSeriesGenerator(new Random(31));
            var manualStandardValue = mpe + 20d;

            var values = generator.GenerateResponseTimeValues(
                standardValue: manualStandardValue,
                mpe: mpe,
                valueCount: 50,
                decimalPlaces: 2,
                configuration: new GenerationConfiguration { DefaultDistribution = "Uniform" });

            Assert.AreEqual(50, values.Count);
            Assert.IsTrue(values.All(value => Math.Abs(value - manualStandardValue) <= expectedMaximumSpread / 2d + 1e-12));
            Assert.IsTrue(values.Max() - values.Min() <= expectedMaximumSpread + 1e-12);
        }

        [TestMethod]
        public void ResponseTimeValuesStayInsideManualStandardRange()
        {
            var generator = new MeasurementSeriesGenerator(new Random(32));

            var values = generator.GenerateResponseTimeValues(
                standardValue: 40,
                mpe: 60,
                valueCount: 50,
                decimalPlaces: 2,
                configuration: new GenerationConfiguration { DefaultDistribution = "Uniform" },
                measurementLowerBound: 20,
                measurementUpperBound: 30);

            Assert.IsTrue(values.All(value => value >= 20 && value <= 30));
            Assert.IsTrue(values.Max() - values.Min() <= 5d + 1e-12);
        }

        [TestMethod]
        public void ResponseTimeValuesUseConfiguredDifferenceControls()
        {
            var generator = new MeasurementSeriesGenerator(new Random(33));
            var configuration = new GenerationConfiguration
            {
                DefaultDistribution = "Uniform",
                ResponseTimeThresholdSeconds = 100,
                ResponseTimeBelowThresholdMaximumDifferenceSeconds = 2,
                ResponseTimeAboveThresholdMaximumDifferenceSeconds = 4
            };

            var belowThreshold = generator.GenerateResponseTimeValues(80, 100, 50, 2, configuration);
            var aboveThreshold = generator.GenerateResponseTimeValues(120, 120, 50, 2, configuration);

            Assert.IsTrue(belowThreshold.Max() - belowThreshold.Min() <= 2d + 1e-12);
            Assert.IsTrue(aboveThreshold.Max() - aboveThreshold.Min() <= 4d + 1e-12);
        }
    }
}
