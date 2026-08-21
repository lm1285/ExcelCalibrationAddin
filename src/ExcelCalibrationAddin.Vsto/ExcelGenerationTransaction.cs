using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class ExcelGenerationTransaction : IDisposable
    {
        private readonly List<RangeSnapshot> _snapshots = new List<RangeSnapshot>();
        private bool _committed;

        public ExcelGenerationTransaction(Excel.Workbook workbook, IEnumerable<MeasurementRule> rules)
        {
            if (workbook == null)
            {
                return;
            }

            foreach (var range in (rules ?? Enumerable.Empty<MeasurementRule>())
                         .Where(rule => rule?.IsEnabled == true && rule.TargetRange != null)
                         .Select(rule => rule.TargetRange)
                         .GroupBy(BuildKey, StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                var worksheet = workbook.Worksheets[range.SheetName] as Excel.Worksheet;
                var target = worksheet?.Range[
                    worksheet.Cells[range.StartRow, range.StartColumn],
                    worksheet.Cells[range.EndRow, range.EndColumn]];
                if (target != null)
                {
                    _snapshots.Add(new RangeSnapshot(target, target.Formula));
                }
            }
        }

        public void Commit()
        {
            _committed = true;
        }

        public void Dispose()
        {
            if (_committed)
            {
                return;
            }

            foreach (var snapshot in _snapshots)
            {
                try
                {
                    snapshot.Range.Formula = snapshot.Formula;
                }
                catch
                {
                }
            }
        }

        private static string BuildKey(CellRange range)
        {
            return $"{range.SheetName}:{range.StartRow}:{range.StartColumn}:{range.EndRow}:{range.EndColumn}";
        }

        private sealed class RangeSnapshot
        {
            public RangeSnapshot(dynamic range, object formula)
            {
                Range = range;
                Formula = formula;
            }

            public dynamic Range { get; }
            public object Formula { get; }
        }
    }
}
