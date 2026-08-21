using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Recognition;

namespace ExcelCalibrationAddin.Host.Services
{
    public sealed partial class MeasurementRuleStructureAnalyzer
    {
        private static IReadOnlyList<CellReference> ExtractReferences(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return new List<CellReference>();
            }

            return CellReferenceRegex.Matches(formula)
                .Cast<Match>()
                .Select(match => ParseReference(match.Value))
                .Where(reference => reference != null)
                .ToList();
        }

        private static CellReference ParseReference(string value)
        {
            var text = (value ?? string.Empty).Replace("$", string.Empty).ToUpperInvariant();
            var letters = new string(text.TakeWhile(char.IsLetter).ToArray());
            var digits = new string(text.SkipWhile(char.IsLetter).ToArray());
            if (letters.Length == 0 || !int.TryParse(digits, out var row))
            {
                return null;
            }

            var column = 0;
            foreach (var ch in letters)
            {
                column = column * 26 + (ch - 'A' + 1);
            }

            return column <= 0 || row <= 0 ? null : new CellReference(row, column);
        }

        private static bool RangeContains(CellRange range, int row, int column)
        {
            return range != null &&
                row >= range.StartRow &&
                row <= range.EndRow &&
                column >= range.StartColumn &&
                column <= range.EndColumn;
        }

        private static CellRange CloneRange(CellRange range)
        {
            if (range == null)
            {
                return null;
            }

            return new CellRange
            {
                SheetName = range.SheetName,
                StartRow = range.StartRow,
                EndRow = range.EndRow,
                StartColumn = range.StartColumn,
                EndColumn = range.EndColumn
            };
        }

        private sealed class LogicalCell
        {
            public CellMeta Anchor { get; set; }
            public CellRange Range { get; set; }
        }

        private sealed class CellReference
        {
            public CellReference(int row, int column)
            {
                Row = row;
                Column = column;
            }

            public int Row { get; }
            public int Column { get; }
        }
    }
}
