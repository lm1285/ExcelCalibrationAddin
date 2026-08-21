using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropSnapshotProvider : IWorkbookSnapshotProvider, IFingerprintSnapshotProvider
    {
        private const int MaxRowsToScan = 200;
        private const int MaxColumnsToScan = 100;
        private const int MaxCellsToScan = 20000;
        private const int HeaderRowsToInspect = 6;

        private readonly dynamic _workbook;
        private readonly string _requestedSheetName;
        private readonly bool _wasSavedAtCreation;

        public ExcelInteropSnapshotProvider(dynamic workbook, string requestedSheetName = null)
        {
            _workbook = workbook;
            _requestedSheetName = requestedSheetName ?? string.Empty;
            try
            {
                _wasSavedAtCreation = SafeToBool(workbook?.Saved);
            }
            catch
            {
                _wasSavedAtCreation = false;
            }
        }

        public WorkbookSnapshot Capture()
        {
            EnsureCalculated();

            return CaptureCurrentSheet();
        }

        public WorkbookSnapshot CaptureFingerprint()
        {
            EnsureCalculated();
            return CaptureCurrentSheet();
        }

        private WorkbookSnapshot CaptureCurrentSheet()
        {
            var snapshot = new WorkbookSnapshot
            {
                WorkbookName = SafeToString(_workbook.Name)
            };

            var activeSheetName = GetActiveSheetName();
            if (!string.IsNullOrWhiteSpace(activeSheetName))
            {
                RecognitionProgress.Report(12, "正在扫描当前工作表...");
                foreach (var worksheet in _workbook.Worksheets)
                {
                    if (string.Equals(SafeToString(worksheet.Name), activeSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        snapshot.Sheets.Add(CaptureSheet(worksheet));
                        return snapshot;
                    }
                }
            }

            foreach (var worksheet in _workbook.Worksheets)
            {
                snapshot.Sheets.Add(CaptureSheet(worksheet));
            }

            return snapshot;
        }

        public WorkbookSnapshot Capture(IEnumerable<CellRange> ranges)
        {
            EnsureCalculated();

            var snapshot = new WorkbookSnapshot
            {
                WorkbookName = SafeToString(_workbook.Name)
            };

            var rangesBySheet = (ranges ?? Enumerable.Empty<CellRange>())
                .Where(IsValidRange)
                .GroupBy(range => range.SheetName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => MergeRanges(group.ToList()),
                    StringComparer.OrdinalIgnoreCase);

            if (rangesBySheet.Count == 0)
            {
                return Capture();
            }

            foreach (var worksheet in _workbook.Worksheets)
            {
                var sheetName = SafeToString(worksheet.Name);
                List<CellRange> sheetRanges;
                if (!rangesBySheet.TryGetValue(sheetName, out sheetRanges))
                {
                    continue;
                }

                snapshot.Sheets.Add(CaptureSheetRanges(worksheet, sheetRanges));
            }

            return snapshot;
        }

        public string GetActiveSheetName()
        {
            if (!string.IsNullOrWhiteSpace(_requestedSheetName))
            {
                return _requestedSheetName;
            }

            try
            {
                return SafeToString(_workbook.ActiveSheet?.Name);
            }
            catch
            {
                return string.Empty;
            }
        }

    }
}
