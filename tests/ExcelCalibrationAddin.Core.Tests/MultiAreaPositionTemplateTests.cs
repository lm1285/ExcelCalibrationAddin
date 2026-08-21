using System;
using System.Collections.Generic;
using System.IO;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelCalibrationAddin.Core.Tests
{
    [TestClass]
    public sealed class MultiAreaPositionTemplateTests
    {
        [TestMethod]
        public void PositionTemplateKeepsOnlyRelativeAreaGeometry()
        {
            var template = MultiAreaPositionTemplate.Create("报告摘录", new[]
            {
                Area(10, 4, 2, 3),
                Area(20, 8, 1, 2)
            });

            Assert.AreEqual(2, template.Areas.Count);
            Assert.AreEqual(0, template.Areas[0].RowOffset);
            Assert.AreEqual(0, template.Areas[0].ColumnOffset);
            Assert.AreEqual(10, template.Areas[1].RowOffset);
            Assert.AreEqual(4, template.Areas[1].ColumnOffset);

            var resolved = template.Resolve(100, 6);
            Assert.AreEqual(100, resolved[0].StartRow);
            Assert.AreEqual(6, resolved[0].StartColumn);
            Assert.AreEqual(110, resolved[1].StartRow);
            Assert.AreEqual(10, resolved[1].StartColumn);
        }

        [TestMethod]
        public void StoreUpdatesSameNamedTemplateWithoutCreatingDuplicate()
        {
            var directory = Path.Combine(Path.GetTempPath(), "ExcelCalibrationAddin.Tests", Guid.NewGuid().ToString("N"));
            var path = Path.Combine(directory, "multi-area.json");
            var store = new MultiAreaPositionTemplateStore(path);

            var first = store.Save("常用位置", new[] { Area(2, 2, 1, 1) });
            var second = store.Save("常用位置", new[] { Area(4, 4, 2, 2), Area(8, 8, 1, 1) });

            Assert.AreEqual(first.Id, second.Id);
            Assert.AreEqual(1, store.List().Count);
            Assert.AreEqual(2, store.List()[0].Areas.Count);
        }

        [TestMethod]
        public void PositionTemplateRejectsOverlappingAreas()
        {
            Assert.ThrowsException<InvalidOperationException>(() =>
                MultiAreaPositionTemplate.Create("重叠", new List<AbsoluteAreaPosition>
                {
                    Area(1, 1, 3, 3),
                    Area(2, 2, 2, 2)
                }));
        }

        [TestMethod]
        public void StorePersistsSourceIdentityWithoutCellValues()
        {
            var path = Path.Combine(Path.GetTempPath(), "ExcelCalibrationAddin.Tests", Guid.NewGuid().ToString("N"), "multi-area.json");
            var template = MultiAreaPositionTemplate.Create("位置模板", new[] { Area(1, 1, 2, 2) });
            template.SourceWorkbookName = "模板.xlsx";
            template.SourceSheetName = "Sheet1";
            template.SourceAnchorRow = 1;
            template.SourceAnchorColumn = 1;
            var store = new MultiAreaPositionTemplateStore(path);
            store.Save(template);

            var loaded = new MultiAreaPositionTemplateStore(path).List()[0];
            Assert.AreEqual("模板.xlsx", loaded.SourceWorkbookName);
            Assert.AreEqual("Sheet1", loaded.SourceSheetName);
            Assert.IsFalse(File.ReadAllText(path).Contains("Values"));
        }

        private static AbsoluteAreaPosition Area(int row, int column, int rows, int columns)
        {
            return new AbsoluteAreaPosition
            {
                StartRow = row,
                StartColumn = column,
                RowCount = rows,
                ColumnCount = columns
            };
        }
    }
}
