using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class DiagnosticPackageInput
    {
        public string WorkbookName { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public string Fingerprint { get; set; } = string.Empty;
        public string MatchResultJson { get; set; } = string.Empty;
        public string RecognitionResultJson { get; set; } = string.Empty;
        public string SummaryJson { get; set; } = string.Empty;
    }

    public sealed class DiagnosticPackageService
    {
        public Task AppendSummaryAsync(string eventName, object summary)
        {
            return Task.Run(() =>
            {
                try
                {
                    AddinFileLogger.Configure("Diagnostics");
                    TraceSummary(eventName, summary);
                }
                catch
                {
                }
            });
        }

        public string Export(string outputPath, DiagnosticPackageInput input)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("诊断包输出路径不能为空。", nameof(outputPath));
            }

            var fullPath = Path.GetFullPath(outputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? ".");
            if (File.Exists(fullPath))
            {
                File.Delete(fullPath);
            }

            input = input ?? new DiagnosticPackageInput();
            using (var archive = ZipFile.Open(fullPath, ZipArchiveMode.Create))
            {
                AddText(archive, "environment.json", JsonConvert.SerializeObject(new
                {
                    os = Environment.OSVersion.VersionString,
                    runtime = Environment.Version.ToString(),
                    process = AppDomain.CurrentDomain.FriendlyName,
                    created_at = DateTime.UtcNow.ToString("o")
                }, Formatting.Indented));
                AddText(archive, "template-fingerprint.txt", input.Fingerprint ?? string.Empty);
                AddText(archive, "recognition-result.json", input.RecognitionResultJson ?? "{}");
                AddText(archive, "match-result.json", input.MatchResultJson ?? "{}");
                AddText(archive, "diagnostic-summary.json", input.SummaryJson ?? "{}");
                AddLog(archive, "addin.log", AddinFileLogger.LogFilePath);
            }

            return fullPath;
        }

        private static void TraceSummary(string eventName, object summary)
        {
            Trace.WriteLine($"[Diagnostics] {eventName}: {JsonConvert.SerializeObject(summary)}");
        }

        private static void AddText(ZipArchive archive, string name, string content)
        {
            var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
            using (var stream = entry.Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(content ?? string.Empty);
            }
        }

        private static void AddLog(ZipArchive archive, string name, string path)
        {
            if (!File.Exists(path))
            {
                AddText(archive, name, string.Empty);
                return;
            }

            var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
            using (var source = File.OpenRead(path))
            using (var target = entry.Open())
            {
                source.CopyTo(target);
            }
        }
    }
}
