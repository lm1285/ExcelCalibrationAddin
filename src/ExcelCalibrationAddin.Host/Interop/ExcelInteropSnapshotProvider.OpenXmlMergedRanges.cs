using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Interop
{
    public sealed partial class ExcelInteropSnapshotProvider
    {
        private static readonly Regex CellReferenceRegex = new Regex(
            @"^\$?([A-Z]+)\$?(\d+)$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private bool CanUsePersistedMergeLayout()
        {
            return _wasSavedAtCreation;
        }

        private bool TryCapturePersistedMergedRanges(
            dynamic worksheet,
            out List<CellRange> mergedRanges)
        {
            mergedRanges = null;
            try
            {
                string workbookPath = SafeToString(_workbook.FullName);
                string extension = Path.GetExtension(workbookPath);
                if (!File.Exists(workbookPath) ||
                    !new[] { ".xlsx", ".xlsm", ".xltx", ".xltm" }
                        .Contains(extension, StringComparer.OrdinalIgnoreCase))
                {
                    return false;
                }

                using (var stream = new FileStream(
                    workbookPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete))
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, false))
                {
                    var workbookDocument = LoadXml(archive, "xl/workbook.xml");
                    var relationshipsDocument = LoadXml(archive, "xl/_rels/workbook.xml.rels");
                    if (workbookDocument == null || relationshipsDocument == null)
                    {
                        return false;
                    }

                    XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                    XNamespace relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
                    string sheetName = SafeToString(worksheet.Name);
                    var sheet = workbookDocument
                        .Descendants(spreadsheet + "sheet")
                        .FirstOrDefault(item => string.Equals(
                            (string)item.Attribute("name"),
                            sheetName,
                            StringComparison.OrdinalIgnoreCase));
                    var relationshipId = (string)sheet?.Attribute(relationships + "id");
                    if (string.IsNullOrWhiteSpace(relationshipId))
                    {
                        return false;
                    }

                    var relationship = relationshipsDocument
                        .Descendants(packageRelationships + "Relationship")
                        .FirstOrDefault(item => string.Equals(
                            (string)item.Attribute("Id"),
                            relationshipId,
                            StringComparison.Ordinal));
                    var target = (string)relationship?.Attribute("Target");
                    var worksheetPath = ResolveArchivePath("xl/workbook.xml", target);
                    var worksheetDocument = LoadXml(archive, worksheetPath);
                    if (worksheetDocument == null)
                    {
                        return false;
                    }

                    mergedRanges = worksheetDocument
                        .Descendants(spreadsheet + "mergeCell")
                        .Select(item => ParseMergeReference((string)item.Attribute("ref"), sheetName))
                        .Where(range => range != null)
                        .ToList();
                    return true;
                }
            }
            catch
            {
                mergedRanges = null;
                return false;
            }
        }

        private static XDocument LoadXml(ZipArchive archive, string path)
        {
            if (archive == null || string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            var entry = archive.Entries.FirstOrDefault(item => string.Equals(
                item.FullName,
                path,
                StringComparison.OrdinalIgnoreCase));
            if (entry == null)
            {
                return null;
            }

            using (var stream = entry.Open())
            {
                return XDocument.Load(stream);
            }
        }

        private static string ResolveArchivePath(string sourcePath, string target)
        {
            if (string.IsNullOrWhiteSpace(target))
            {
                return string.Empty;
            }

            var normalizedTarget = target.Replace('\\', '/').TrimStart('/');
            var sourceDirectory = sourcePath.Substring(0, sourcePath.LastIndexOf('/') + 1);
            var parts = (normalizedTarget.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
                    ? normalizedTarget
                    : sourceDirectory + normalizedTarget)
                .Split('/');
            var resolved = new List<string>();
            foreach (var part in parts)
            {
                if (string.IsNullOrWhiteSpace(part) || part == ".")
                {
                    continue;
                }
                if (part == "..")
                {
                    if (resolved.Count == 0)
                    {
                        return string.Empty;
                    }
                    resolved.RemoveAt(resolved.Count - 1);
                    continue;
                }
                resolved.Add(part);
            }
            return string.Join("/", resolved);
        }

        private static CellRange ParseMergeReference(string reference, string sheetName)
        {
            var parts = (reference ?? string.Empty).Split(':');
            if (parts.Length < 1 || parts.Length > 2 ||
                !TryParseCellReference(parts[0], out var startRow, out var startColumn) ||
                !TryParseCellReference(parts.Length == 2 ? parts[1] : parts[0], out var endRow, out var endColumn))
            {
                return null;
            }

            return new CellRange
            {
                SheetName = sheetName,
                StartRow = Math.Min(startRow, endRow),
                StartColumn = Math.Min(startColumn, endColumn),
                EndRow = Math.Max(startRow, endRow),
                EndColumn = Math.Max(startColumn, endColumn)
            };
        }

        private static bool TryParseCellReference(string reference, out int row, out int column)
        {
            row = 0;
            column = 0;
            var match = CellReferenceRegex.Match((reference ?? string.Empty).Trim());
            if (!match.Success || !int.TryParse(match.Groups[2].Value, out row) || row <= 0)
            {
                return false;
            }

            foreach (var character in match.Groups[1].Value.ToUpperInvariant())
            {
                column = checked(column * 26 + character - 'A' + 1);
            }
            return column > 0;
        }
    }
}
