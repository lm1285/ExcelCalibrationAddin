using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace ExcelCalibrationAddin.Core.Services
{
    public static class AddinFileLogger
    {
        private const long MaxLogBytes = 10 * 1024 * 1024;
        private static readonly object SyncRoot = new object();
        private static bool _configured;

        public static string LogFilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelCalibrationAddin",
                    "addin.log");
            }
        }

        public static void Configure(string processName)
        {
            lock (SyncRoot)
            {
                if (_configured)
                {
                    return;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath) ?? ".");
                Trace.Listeners.Add(new SharedFileTraceListener(LogFilePath, processName));
                Trace.AutoFlush = true;
                _configured = true;
            }

            Trace.WriteLine($"[Logger] File logging enabled. Path={LogFilePath}");
        }

        private sealed class SharedFileTraceListener : TraceListener
        {
            private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);
            private readonly string _path;
            private readonly string _processName;

            public SharedFileTraceListener(string path, string processName)
            {
                _path = path;
                _processName = string.IsNullOrWhiteSpace(processName) ? AppDomain.CurrentDomain.FriendlyName : processName;
            }

            public override void Write(string message)
            {
                Append(message, includePrefix: false, includeNewLine: false);
            }

            public override void WriteLine(string message)
            {
                Append(message, includePrefix: true, includeNewLine: true);
            }

            private void Append(string message, bool includePrefix, bool includeNewLine)
            {
                try
                {
                    lock (SyncRoot)
                    {
                        RotateIfNeeded();
                        using (var stream = new FileStream(_path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                        using (var writer = new StreamWriter(stream, Utf8NoBom))
                        {
                            if (includePrefix)
                            {
                                writer.Write($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{_processName}] ");
                            }

                            writer.Write(message ?? string.Empty);
                            if (includeNewLine)
                            {
                                writer.WriteLine();
                            }
                        }
                    }
                }
                catch
                {
                }
            }

            private void RotateIfNeeded()
            {
                try
                {
                    var file = new FileInfo(_path);
                    if (!file.Exists || file.Length < MaxLogBytes)
                    {
                        return;
                    }

                    var archivePath = _path + ".1";
                    if (File.Exists(archivePath))
                    {
                        File.Delete(archivePath);
                    }

                    File.Move(_path, archivePath);
                }
                catch
                {
                }
            }
        }
    }
}
