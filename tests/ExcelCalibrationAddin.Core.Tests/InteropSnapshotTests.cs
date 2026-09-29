using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelCalibrationAddin.Core.Tests
{
    [TestClass]
    public sealed class InteropSnapshotTests
    {
        [TestMethod]
        public void UnsavedPrecisionChangePreservesPersistedFingerprintAndBlankMerges()
        {
            using (var workbook = new FakeWorkbook())
            {
                var saved = new ExcelInteropSnapshotProvider(workbook).CaptureFingerprint();
                workbook.Saved = false;
                workbook.ActiveSheet.MeasurementFormat = "0.0000";
                var current = new ExcelInteropSnapshotProvider(workbook).CaptureFingerprint();

                Assert.IsTrue(current.Sheets[0].Cells.Single(cell => cell.Row == 10 && cell.Column == 10).IsMerged);
                Assert.IsTrue(current.Sheets[0].Cells.Single(cell => cell.Row == 12 && cell.Column == 20).IsMerged);
                Assert.AreEqual("0.0000", current.Sheets[0].Cells.Single(cell => cell.Row == 8 && cell.Column == 2).NumberFormat);
                var builder = new TemplateFingerprintBuilder();
                Assert.AreEqual(builder.Build(saved).ExactFingerprint, builder.Build(current).ExactFingerprint);
                Assert.AreEqual(builder.Build(saved).FuzzyFingerprint, builder.Build(current).FuzzyFingerprint);
                Assert.IsTrue(workbook.ActiveSheet.MergeStateReads < 200,
                    "Rows without merges should be skipped instead of querying all 1200 cells.");
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void UnsavedMergeEditUsesLiveLayout(bool captureSelectedRanges)
        {
            using (var workbook = new FakeWorkbook())
            {
                var saved = new ExcelInteropSnapshotProvider(workbook).CaptureFingerprint();
                workbook.Saved = false;
                workbook.ActiveSheet.Merges.Add(new CellRange
                {
                    SheetName = "Sheet1", StartRow = 15, EndRow = 15, StartColumn = 3, EndColumn = 5
                });
                var provider = new ExcelInteropSnapshotProvider(workbook);
                var current = captureSelectedRanges
                    ? provider.Capture(new[] { new CellRange
                        { SheetName = "Sheet1", StartRow = 1, EndRow = 40, StartColumn = 1, EndColumn = 30 } })
                    : provider.CaptureFingerprint();

                Assert.IsTrue(current.Sheets[0].Cells.Single(cell => cell.Row == 15 && cell.Column == 4).IsMerged);
                var builder = new TemplateFingerprintBuilder();
                Assert.AreNotEqual(builder.Build(saved).ExactFingerprint, builder.Build(current).ExactFingerprint);
            }
        }

        public sealed class FakeWorkbook : IDisposable
        {
            public FakeWorkbook()
            {
                FullName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
                using (var stream = File.Create(FullName))
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
                {
                    WriteEntry(archive, "xl/workbook.xml",
                        "<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='Sheet1' sheetId='1' r:id='rId1'/></sheets></workbook>");
                    WriteEntry(archive, "xl/_rels/workbook.xml.rels",
                        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml'/></Relationships>");
                    WriteEntry(archive, "xl/worksheets/sheet1.xml",
                        "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><mergeCells><mergeCell ref='J10:L10'/><mergeCell ref='T11:T13'/></mergeCells></worksheet>");
                }
            }

            public bool Saved { get; set; } = true;
            public string Name => "Regression.xlsx";
            public string FullName { get; }
            public FakeSheet ActiveSheet { get; } = new FakeSheet();
            public FakeSheet[] Worksheets => new[] { ActiveSheet };
            public object[] Names => Array.Empty<object>();
            public FakeApplication Application { get; } = new FakeApplication();
            public void Dispose() => File.Delete(FullName);

            private static void WriteEntry(ZipArchive archive, string path, string xml)
            {
                using (var writer = new StreamWriter(archive.CreateEntry(path).Open()))
                {
                    writer.Write(xml);
                }
            }
        }

        public sealed class FakeApplication
        {
            public int CalculationState => 0;
        }

        public sealed class FakePageSetup
        {
            public string PrintArea => string.Empty;
        }

        public sealed class FakeSheet
        {
            public FakeSheet()
            {
                Cells = new FakeCells(this);
                Range = new FakeRanges(this);
            }

            public string Name => "Sheet1";
            public string MeasurementFormat { get; set; } = "0.0";
            public int MergeStateReads { get; set; }
            public FakeCells Cells { get; }
            public FakeRanges Range { get; }
            public FakePageSetup PageSetup { get; } = new FakePageSetup();
            public FakeRange UsedRange => new FakeRange(this, 1, 1, 40, 30);
            public List<CellRange> Merges { get; } = new List<CellRange>
            {
                new CellRange { SheetName = "Sheet1", StartRow = 10, EndRow = 10, StartColumn = 10, EndColumn = 12 },
                new CellRange { SheetName = "Sheet1", StartRow = 11, EndRow = 13, StartColumn = 20, EndColumn = 20 }
            };
        }

        public sealed class FakeCells
        {
            private readonly FakeSheet _sheet;
            public FakeCells(FakeSheet sheet) => _sheet = sheet;
            public FakeRange this[int row, int column] => new FakeRange(_sheet, row, column, row, column);
        }

        public sealed class FakeRanges
        {
            private readonly FakeSheet _sheet;
            public FakeRanges(FakeSheet sheet) => _sheet = sheet;
            public FakeRange this[FakeRange first, FakeRange last] =>
                new FakeRange(_sheet, first.Row, first.Column, last.EndRow, last.EndColumn);
        }

        public sealed class FakeCount
        {
            public int Count { get; set; }
        }

        public sealed class FakeRange
        {
            private readonly FakeSheet _sheet;
            public FakeRange(FakeSheet sheet, int row, int column, int endRow, int endColumn)
            {
                _sheet = sheet;
                Row = row;
                Column = column;
                EndRow = endRow;
                EndColumn = endColumn;
            }

            public int Row { get; }
            public int Column { get; }
            public int EndRow { get; }
            public int EndColumn { get; }
            public FakeCount Rows => new FakeCount { Count = EndRow - Row + 1 };
            public FakeCount Columns => new FakeCount { Count = EndColumn - Column + 1 };
            public object Value2 => ReadValues(false);
            public object Formula => Value2;
            public object FormulaR1C1 => Value2;
            public object NumberFormat => ReadValues(true);
            public object Text => null;

            public bool? MergeCells
            {
                get
                {
                    _sheet.MergeStateReads++;
                    if (!_sheet.Merges.Any(merge => merge.StartRow <= EndRow && merge.EndRow >= Row &&
                        merge.StartColumn <= EndColumn && merge.EndColumn >= Column))
                    {
                        return false;
                    }
                    return Rows.Count == 1 && Columns.Count == 1 ? true : (bool?)null;
                }
            }

            public FakeRange MergeArea
            {
                get
                {
                    var merge = _sheet.Merges.Single(item => item.StartRow <= Row && item.EndRow >= Row &&
                        item.StartColumn <= Column && item.EndColumn >= Column);
                    return new FakeRange(_sheet, merge.StartRow, merge.StartColumn, merge.EndRow, merge.EndColumn);
                }
            }

            private object ReadValues(bool formats)
            {
                var values = new object[Rows.Count + 1, Columns.Count + 1];
                for (var row = 1; row <= Rows.Count; row++)
                {
                    for (var column = 1; column <= Columns.Count; column++)
                    {
                        var sheetRow = Row + row - 1;
                        var sheetColumn = Column + column - 1;
                        values[row, column] = formats
                            ? (sheetRow == 8 && sheetColumn == 2 ? _sheet.MeasurementFormat : "General")
                            : sheetRow == 1 && sheetColumn == 1 ? (object)"Calibration"
                            : sheetRow == 8 && sheetColumn == 2 ? 1.2345 : null;
                    }
                }
                return Rows.Count == 1 && Columns.Count == 1 ? values[1, 1] : values;
            }
        }
    }
}
