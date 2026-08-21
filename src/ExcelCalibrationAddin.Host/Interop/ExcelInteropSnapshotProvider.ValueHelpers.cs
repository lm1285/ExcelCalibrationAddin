using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.UseCases;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropSnapshotProvider
    {
        private static string SafeToString(dynamic value)
        {
            try
            {
                return value == null ? string.Empty : Convert.ToString(value);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ResolveCellText(string displayText, string rawValueText)
        {
            var display = displayText ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(display) && !LooksLikeUnavailableDisplayText(display))
            {
                return display;
            }

            return rawValueText ?? string.Empty;
        }

        private static bool LooksLikeUnavailableDisplayText(string text)
        {
            var value = (text ?? string.Empty).Trim();
            return value.Length > 0 && value.All(ch => ch == '#');
        }

        private static string NormalizeFormula(dynamic value)
        {
            var formula = SafeToString(value).Trim();
            return formula.StartsWith("=", StringComparison.Ordinal) ? formula : string.Empty;
        }

        private static int SafeToInt(dynamic value)
        {
            try
            {
                return value == null ? 0 : Convert.ToInt32(value);
            }
            catch
            {
                return 0;
            }
        }

        private static bool SafeToBool(dynamic value)
        {
            try
            {
                return value != null && Convert.ToBoolean(value);
            }
            catch
            {
                return false;
            }
        }

        private sealed class ScanArea
        {
            public int StartRow { get; set; }
            public int StartColumn { get; set; }
            public int EndRow { get; set; }
            public int EndColumn { get; set; }

            public int RowCount => Math.Max(0, EndRow - StartRow + 1);
            public int ColumnCount => Math.Max(0, EndColumn - StartColumn + 1);

            public string Address =>
                $"{ExcelAddressHelper.ToColumnName(StartColumn)}{StartRow}:{ExcelAddressHelper.ToColumnName(EndColumn)}{EndRow}";
        }

        private sealed class CellAddress
        {
            public CellAddress(int row, int column)
            {
                Row = row;
                Column = column;
            }

            public int Row { get; }
            public int Column { get; }
        }

    }
}
