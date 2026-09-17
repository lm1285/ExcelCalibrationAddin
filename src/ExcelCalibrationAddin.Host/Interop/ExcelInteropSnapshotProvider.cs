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
                WorkbookName = SafeToString(_workbook.Name),
                NamedRanges = CaptureNamedRanges()
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
                WorkbookName = SafeToString(_workbook.Name),
                NamedRanges = CaptureNamedRanges()
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

        private List<NamedRangeDefinition> CaptureNamedRanges()
        {
            var result = new List<NamedRangeDefinition>();
            try
            {
                foreach (var definedName in _workbook.Names)
                {
                    try
                    {
                        dynamic refersToRange = definedName.RefersToRange;
                        dynamic worksheet = refersToRange.Worksheet;
                        var workbookName = SafeToString(worksheet.Parent?.Name);
                        if (!string.Equals(workbookName, SafeToString(_workbook.Name), StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        var row = SafeToInt(refersToRange.Row);
                        var column = SafeToInt(refersToRange.Column);
                        var rowCount = SafeToInt(refersToRange.Rows.Count);
                        var columnCount = SafeToInt(refersToRange.Columns.Count);
                        if (row <= 0 || column <= 0 || rowCount <= 0 || columnCount <= 0)
                        {
                            continue;
                        }

                        var name = NormalizeDefinedName(SafeToString(definedName.Name));
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            continue;
                        }

                        result.Add(new NamedRangeDefinition
                        {
                            Name = name,
                            Range = new CellRange
                            {
                                SheetName = SafeToString(worksheet.Name),
                                StartRow = row,
                                StartColumn = column,
                                EndRow = row + rowCount - 1,
                                EndColumn = column + columnCount - 1
                            }
                        });
                    }
                    catch
                    {
                        // Constants, formulas and external names do not resolve to a local range.
                    }
                }
            }
            catch
            {
                // A workbook without accessible names remains valid.
            }

            return result
                .GroupBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();
        }

        private static string NormalizeDefinedName(string value)
        {
            var name = (value ?? string.Empty).Trim().Trim('\'');
            var separator = name.LastIndexOf('!');
            return (separator >= 0 ? name.Substring(separator + 1) : name).Trim().Trim('\'');
        }

    }
}
