using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private static void WriteGenerationPerformanceLog(string message)
        {
            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelCalibrationAddin");
                Directory.CreateDirectory(directory);
                var path = Path.Combine(directory, "generation-performance.log");
                File.AppendAllText(path, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}{Environment.NewLine}");
            }
            catch
            {
            }
        }

        private void RefreshRandomRangeSummary(IEnumerable<MeasurementRule> rules, GenerationConfiguration configuration = null)
        {
            try
            {
                Globals.Ribbons?.CalibrationRibbon?.UpdateRandomRangeSummary(configuration ?? LoadGenerationConfiguration(), rules);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Refresh random range summary failed: {ex}");
            }
        }

        private void ReportRecognitionProgress(int percent, string message)
        {
            var safePercent = Math.Max(0, Math.Min(100, percent));
            var safeMessage = string.IsNullOrWhiteSpace(message) ? "\u6b63\u5728\u8bc6\u522b..." : message;
            Trace.WriteLine($"[VSTO] Progress {safePercent}% - {safeMessage}");
            if (!ShouldUpdateProgressUi(safePercent, safeMessage))
            {
                return;
            }

            _taskPaneControl?.SetRecognitionProgress(safeMessage, safePercent, true);
            Application.StatusBar = $"\u6821\u51c6\u52a9\u624b\uff1a{safeMessage} ({safePercent}%)";
        }

        private bool ShouldUpdateProgressUi(int percent, string message)
        {
            var now = DateTime.UtcNow;
            var isFinal = percent >= 100;
            var messageChanged = !string.Equals(message, _lastProgressMessage, StringComparison.Ordinal);
            var percentChangedEnough = Math.Abs(percent - _lastProgressPercent) >= 3;
            var intervalElapsed = (now - _lastProgressUiUpdateUtc).TotalMilliseconds >= 150;

            if (!isFinal && !messageChanged && !percentChangedEnough && !intervalElapsed)
            {
                return false;
            }

            _lastProgressUiUpdateUtc = now;
            _lastProgressPercent = percent;
            _lastProgressMessage = message;
            return true;
        }
    }
}
