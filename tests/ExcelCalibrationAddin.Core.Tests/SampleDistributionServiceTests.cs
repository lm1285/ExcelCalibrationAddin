using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Core.Repositories;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelCalibrationAddin.Core.Tests
{
    [TestClass]
    public sealed class SampleDistributionServiceTests
    {
        [TestMethod]
        public void GeneratesRoundedValuesWithinExpandedSampleRange()
        {
            var service = new SampleDistributionService(new System.Random(7));
            var points = new[] { new SampleDataPoint { CalibrationItemName = "示值误差", StandardValue = 10, DecimalPlaces = 2, MeasurementValues = new List<double> { 9.91, 10.02, 10.08, 9.98 } } };
            Assert.IsTrue(service.TryGenerate(points, "示值误差", 10, 12, out var values, out var decimals));
            Assert.AreEqual(2, decimals);
            Assert.AreEqual(12, values.Count);
            Assert.IsTrue(values.All(value => value >= 9.89 && value <= 10.10));
            Assert.IsTrue(values.All(value => System.Math.Abs(value - System.Math.Round(value, 2)) < 1e-12));
        }

        [TestMethod]
        public void RejectsInsufficientSamples()
        {
            var service = new SampleDistributionService(new System.Random(1));
            Assert.IsFalse(service.TryGenerate(new[] { new SampleDataPoint { MeasurementValues = new List<double> { 1, 2 } } }, null, null, 3, out _, out _));
        }

        [TestMethod]
        public void RepositoryKeepsVersionsAndDeletesChildren()
        {
            var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "sample-data-" + System.Guid.NewGuid().ToString("N") + ".sqlite");
            try
            {
                var repository = new LocalTemplateRuleCacheRepository(path);
                repository.Initialize();
                var versionId = repository.SaveSampleDataVersion("fp", new[]
                {
                    new TemplateSampleData
                    {
                        CalibrationItemName = "项目",
                        CalibrationItemKey = "item",
                        Points = new List<SampleDataPoint> { new SampleDataPoint { PointIndex = 1, MeasurementValues = new List<double> { 1, 2, 3 }, DecimalPlaces = 1 } }
                    }
                });
                Assert.AreEqual(1, repository.ListSampleDataVersions("fp").Count);
                Assert.AreEqual(1, repository.ListLatestSampleDataPoints("fp").Count);
                Assert.IsTrue(repository.DeleteSampleDataVersion(versionId));
                Assert.AreEqual(0, repository.ListLatestSampleDataPoints("fp").Count);
            }
            finally
            {
                if (System.IO.File.Exists(path)) System.IO.File.Delete(path);
            }
        }
    }
}
