using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class TemplateFingerprintBuilder
    {
        private static readonly Regex NumericValueWithUnitRegex = new Regex(
            @"^[+-]?\d+(\.\d+)?([eE][+-]?\d+)?\s*[^\d\s]+(\s*[^\d\s]+)*$",
            RegexOptions.Compiled);

        public TemplateFingerprint Build(WorkbookSnapshot workbook)
        {
            if (workbook == null)
            {
                return new TemplateFingerprint();
            }

            var sheetNames = workbook.Sheets
                .Select(sheet => NormalizeHeader(sheet.Name))
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .Distinct()
                .ToList();

            var headerTexts = workbook.Sheets
                .SelectMany(sheet => sheet.Headers.Select(header => NormalizeStructuralHeader(header.FullText)))
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .Distinct()
                .ToList();

            var title = ResolveTitle(workbook);
            var structureSource = string.Join("|", workbook.Sheets.Select(BuildStructureSource));
            var exactSource = string.Join("|",
                workbook.Sheets.Select(sheet =>
                    $"{BuildNormalizedUsedRangeShape(sheet)}:{BuildNormalizedHeaderSource(sheet)}")) +
                ":" + structureSource;

            var fuzzySource = string.Join("|",
                workbook.Sheets.Select(sheet =>
                    $"{sheet.Headers.Count}:{sheet.Cells.Count}:{sheet.Cells.Count(cell => cell.IsMerged)}"));

            return new TemplateFingerprint
            {
                ExactFingerprint = Hash(exactSource),
                FuzzyFingerprint = Hash(fuzzySource),
                StructureSignature = structureSource,
                Summary = exactSource,
                SheetNames = sheetNames,
                Title = title,
                HeaderTexts = headerTexts
            };
        }

        private static string BuildStructureSource(SheetSnapshot sheet)
        {
            if (sheet == null)
            {
                return string.Empty;
            }

            var originRow = GetOriginRow(sheet);
            var originColumn = GetOriginColumn(sheet);
            var cells = (sheet.Cells ?? new List<CellMeta>())
                .Where(cell => cell != null)
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .Select(cell => string.Join("~", new[]
                {
                    (cell.Row - originRow + 1).ToString(),
                    (cell.Column - originColumn + 1).ToString(),
                    string.IsNullOrWhiteSpace(cell.Formula) ? NormalizeStructuralText(cell.Text) : string.Empty,
                    NormalizeNumberFormat(cell.NumberFormat),
                    NormalizeMergeRange(cell.MergeRange, originRow, originColumn)
                }))
                .Where(value => value.Split('~').Skip(2).Any(item => !string.IsNullOrWhiteSpace(item)))
                .ToList();

            return $"{NormalizeHeader(sheet.Name)}[{string.Join(";", cells)}]";
        }

        private static string BuildNormalizedUsedRangeShape(SheetSnapshot sheet)
        {
            var cells = GetStructuralCells(sheet);
            if (cells.Count == 0)
            {
                return string.Empty;
            }

            return $"1,1:{cells.Max(cell => cell.Row) - GetOriginRow(sheet) + 1},{cells.Max(cell => cell.Column) - GetOriginColumn(sheet) + 1}";
        }

        private static string BuildNormalizedHeaderSource(SheetSnapshot sheet)
        {
            var originColumn = GetOriginColumn(sheet);
            return string.Join(",", (sheet?.Headers ?? new List<HeaderPath>())
                .Where(header => header != null)
                .Select(header => new
                {
                    RelativeColumn = header.Column - originColumn + 1,
                    Text = NormalizeStructuralHeader(header.FullText)
                })
                .Where(header => !string.IsNullOrWhiteSpace(header.Text))
                .Select(header => $"{header.RelativeColumn}:{header.Text}"));
        }

        private static int GetOriginRow(SheetSnapshot sheet)
        {
            return GetStructuralCells(sheet)
                .Where(cell => cell.Row > 0)
                .Select(cell => cell.Row)
                .DefaultIfEmpty(1)
                .Min();
        }

        private static int GetOriginColumn(SheetSnapshot sheet)
        {
            return GetStructuralCells(sheet)
                .Where(cell => cell.Column > 0)
                .Select(cell => cell.Column)
                .DefaultIfEmpty(1)
                .Min();
        }

        private static List<CellMeta> GetStructuralCells(SheetSnapshot sheet)
        {
            return (sheet?.Cells ?? new List<CellMeta>())
                .Where(cell => cell != null && HasStructuralSignal(cell))
                .ToList();
        }

        private static bool HasStructuralSignal(CellMeta cell)
        {
            return cell != null &&
                (cell.MergeRange != null ||
                 !string.IsNullOrWhiteSpace(cell.Formula) ||
                 !string.IsNullOrWhiteSpace(NormalizeStructuralText(cell.Text)) ||
                 !string.IsNullOrWhiteSpace(NormalizeNumberFormat(cell.NumberFormat)));
        }

        private static string NormalizeMergeRange(CellRange range, int originRow, int originColumn)
        {
            if (range == null)
            {
                return string.Empty;
            }

            return $"{range.StartRow - originRow + 1},{range.StartColumn - originColumn + 1}-{range.EndRow - originRow + 1},{range.EndColumn - originColumn + 1}";
        }

        private static string NormalizeStructuralText(string value)
        {
            var text = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text) || IsDataValue(text))
            {
                return string.Empty;
            }

            return string.Join("", text.Where(ch => !char.IsWhiteSpace(ch)));
        }

        private static string NormalizeNumberFormat(string value)
        {
            var normalized = NormalizeStructuralText(value);
            return string.Equals(normalized, "GENERAL", System.StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : normalized;
        }

        private static bool IsOnlyNumeric(string value)
        {
            var normalized = (value ?? string.Empty).Trim().Replace(",", "");
            double number;
            return double.TryParse(normalized, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out number);
        }

        private static bool IsDataValue(string value)
        {
            var text = (value ?? string.Empty).Trim();
            if (IsOnlyNumeric(text))
            {
                return true;
            }

            if (text.IndexOf('\u00B1') >= 0 ||
                text.IndexOf('<') >= 0 ||
                text.IndexOf('>') >= 0 ||
                text.IndexOf('\u2264') >= 0 ||
                text.IndexOf('\u2265') >= 0)
            {
                return false;
            }

            return NumericValueWithUnitRegex.IsMatch(text) &&
                text.IndexOf("%FS", System.StringComparison.OrdinalIgnoreCase) < 0;
        }

        private static string ResolveTitle(WorkbookSnapshot workbook)
        {
            foreach (var sheet in workbook.Sheets)
            {
                foreach (var cell in sheet.Cells.OrderBy(item => item.Row).ThenBy(item => item.Column))
                {
                    var text = (cell.Text ?? string.Empty).Trim();
                    if (cell.Row > 8 || cell.Column > 8)
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(text) || text.Length < 2 || IsOnlyNumeric(text))
                    {
                        continue;
                    }

                    return text;
                }
            }

            return string.Empty;
        }

        private static string NormalizeHeader(string value)
        {
            return string.Join("/",
                (value ?? string.Empty)
                    .Split('/')
                    .Select(item => item.Trim())
                    .Where(item => !string.IsNullOrWhiteSpace(item))
                    .Distinct());
        }

        private static string NormalizeStructuralHeader(string value)
        {
            return string.Join("/",
                (value ?? string.Empty)
                    .Split('/')
                    .Select(NormalizeStructuralText)
                    .Where(item => !string.IsNullOrWhiteSpace(item))
                    .Distinct());
        }

        private static string Hash(string content)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(content ?? string.Empty));
                return string.Concat(bytes.Select(item => item.ToString("x2")));
            }
        }
    }
}
