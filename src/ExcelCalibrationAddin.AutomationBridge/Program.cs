using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelCalibrationAddin.AutomationBridge
{
    internal static class Program
    {
        private const string AddinProgId = "ExcelStandaloneComAddin.Vsto";
        private const string ApiPrefix = "/api/yingdao";
        private static readonly HttpClient Client = new HttpClient { Timeout = TimeSpan.FromSeconds(120) };
        private static BridgeOptions _options;
        private static HttpListener _listener;
        private static readonly object LogLock = new object();

        private static int Main(string[] args)
        {
            _options = BridgeOptions.Parse(args);
            try
            {
                _listener = new HttpListener();
                _listener.Prefixes.Add("http://127.0.0.1:" + _options.PublicPort + "/");
                _listener.Start();
                WriteLog("Bridge listening on port " + _options.PublicPort + ". Inner port=" + _options.InnerPort);
                _ = MonitorExcelAsync();

                while (_listener.IsListening)
                {
                    var context = _listener.GetContext();
                    _ = HandleRequestAsync(context);
                }

                return 0;
            }
            catch (HttpListenerException ex)
            {
                WriteLog("Bridge could not listen: " + ex.Message);
                return 0;
            }
            catch (Exception ex)
            {
                WriteLog("Bridge fatal error: " + ex);
                return 1;
            }
        }

        private static async Task MonitorExcelAsync()
        {
            while (true)
            {
                await Task.Delay(_options.ExcelProcessId > 0 ? 1000 : 500).ConfigureAwait(false);
                if (_options.ExcelProcessId <= 0)
                {
                    // The bridge is intentionally independent from VSTO. When Excel is
                    // opened normally, keep the COM add-in connected before the first API call.
                    if (Process.GetProcessesByName("EXCEL").Length > 0 &&
                        !await IsInnerServiceHealthyAsync().ConfigureAwait(false))
                    {
                        TryConnectAddin(false);
                    }
                    continue;
                }
                try
                {
                    var process = Process.GetProcessById(_options.ExcelProcessId);
                    if (!process.HasExited)
                    {
                        continue;
                    }
                }
                catch (ArgumentException)
                {
                }

                WriteLog("Excel exited. Bridge stopping.");
                Environment.Exit(0);
            }
        }

        private static async Task HandleRequestAsync(HttpListenerContext context)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(context.Request.Headers["Origin"]))
                {
                    WriteJson(context.Response, HttpStatusCode.Forbidden, "{\"ok\":false,\"error\":\"Browser requests are not allowed.\"}");
                    return;
                }

                if (!IsAuthorized(context.Request))
                {
                    WriteJson(context.Response, HttpStatusCode.Unauthorized,
                        "{\"ok\":false,\"error\":\"Invalid automation token.\"}");
                    return;
                }

                var route = NormalizeRoute(context.Request.Url?.AbsolutePath);
                if (context.Request.HttpMethod.Equals("GET", StringComparison.OrdinalIgnoreCase) && route == "/health")
                {
                    var ready = await EnsureInnerServiceAsync().ConfigureAwait(false);
                    WriteJson(context.Response, ready ? HttpStatusCode.OK : HttpStatusCode.ServiceUnavailable,
                        ready
                            ? "{\"ok\":true,\"service\":\"ExcelCalibrationAddin.YingdaoBridge\",\"port\":" + _options.PublicPort + "}"
                            : "{\"ok\":false,\"error\":\"Excel add-in is unavailable.\"}");
                    return;
                }

                if (route != "/status" && route != "/generate")
                {
                    WriteJson(context.Response, HttpStatusCode.NotFound, "{\"ok\":false,\"error\":\"Unknown endpoint.\"}");
                    return;
                }

                if (!await EnsureInnerServiceAsync().ConfigureAwait(false))
                {
                    WriteJson(context.Response, HttpStatusCode.ServiceUnavailable,
                        "{\"ok\":false,\"error\":\"Excel add-in could not be reconnected.\"}");
                    return;
                }

                await ForwardRequestAsync(context, route).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                WriteLog("Request failed: " + ex);
                WriteJson(context.Response, HttpStatusCode.InternalServerError,
                    "{\"ok\":false,\"error\":\"Automation bridge request failed.\"}");
            }
        }

        private static async Task<bool> EnsureInnerServiceAsync()
        {
            if (await IsInnerServiceHealthyAsync().ConfigureAwait(false))
            {
                return true;
            }

            if (!TryConnectAddin(false))
            {
                return false;
            }

            for (var attempt = 0; attempt < 10; attempt++)
            {
                await Task.Delay(100).ConfigureAwait(false);
                if (await IsInnerServiceHealthyAsync().ConfigureAwait(false))
                {
                    return true;
                }
            }

            TryConnectAddin(true);
            for (var attempt = 0; attempt < 10; attempt++)
            {
                await Task.Delay(100).ConfigureAwait(false);
                if (await IsInnerServiceHealthyAsync().ConfigureAwait(false))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryConnectAddin(bool reconnect)
        {
            try
            {
                dynamic excel = Marshal.GetActiveObject("Excel.Application");
                dynamic addin = excel.COMAddIns.Item(AddinProgId);
                if (reconnect)
                {
                    addin.Connect = false;
                }
                addin.Connect = true;
                WriteLog("Requested Excel add-in connection. Reconnect=" + reconnect);
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Could not connect Excel add-in: " + ex.Message);
                return false;
            }
        }

        private static async Task<bool> IsInnerServiceHealthyAsync()
        {
            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, InnerUrl("/health")))
                {
                    if (!string.IsNullOrWhiteSpace(_options.Token))
                    {
                        request.Headers.TryAddWithoutValidation("X-Excel-Calibration-Token", _options.Token);
                    }

                    using (var response = await Client.SendAsync(request).ConfigureAwait(false))
                    {
                        return response.IsSuccessStatusCode;
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        private static async Task ForwardRequestAsync(HttpListenerContext context, string route)
        {
            var request = context.Request;
            var body = await ReadBodyAsync(request).ConfigureAwait(false);
            using (var forward = new HttpRequestMessage(new HttpMethod(request.HttpMethod), InnerUrl(route)))
            {
                if (body.Length > 0 || request.HttpMethod.Equals("POST", StringComparison.OrdinalIgnoreCase))
                {
                    forward.Content = new ByteArrayContent(body);
                    forward.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(
                        string.IsNullOrWhiteSpace(request.ContentType) ? "application/json" : request.ContentType);
                }

                var token = string.IsNullOrWhiteSpace(request.Headers["X-Excel-Calibration-Token"])
                    ? _options.Token
                    : request.Headers["X-Excel-Calibration-Token"];
                if (!string.IsNullOrWhiteSpace(token))
                {
                    forward.Headers.TryAddWithoutValidation("X-Excel-Calibration-Token", token);
                }

                using (var response = await Client.SendAsync(forward).ConfigureAwait(false))
                {
                    var responseBody = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    context.Response.StatusCode = (int)response.StatusCode;
                    context.Response.ContentType = response.Content.Headers.ContentType?.ToString() ?? "application/json; charset=utf-8";
                    context.Response.ContentLength64 = responseBody.Length;
                    await context.Response.OutputStream.WriteAsync(responseBody, 0, responseBody.Length).ConfigureAwait(false);
                    context.Response.Close();
                }
            }
        }

        private static async Task<byte[]> ReadBodyAsync(HttpListenerRequest request)
        {
            using (var buffer = new MemoryStream())
            {
                await request.InputStream.CopyToAsync(buffer).ConfigureAwait(false);
                return buffer.ToArray();
            }
        }

        private static string NormalizeRoute(string path)
        {
            var route = (path ?? string.Empty).TrimEnd('/');
            if (route.StartsWith(ApiPrefix, StringComparison.OrdinalIgnoreCase))
            {
                route = route.Substring(ApiPrefix.Length);
            }
            return string.IsNullOrWhiteSpace(route) ? "/" : route;
        }

        private static string InnerUrl(string route)
        {
            return "http://127.0.0.1:" + _options.InnerPort + ApiPrefix + route;
        }

        private static bool IsAuthorized(HttpListenerRequest request)
        {
            return string.IsNullOrWhiteSpace(_options.Token) ||
                string.Equals(request.Headers["X-Excel-Calibration-Token"], _options.Token, StringComparison.Ordinal);
        }

        private static void WriteJson(HttpListenerResponse response, HttpStatusCode status, string body)
        {
            try
            {
                var bytes = Encoding.UTF8.GetBytes(body);
                response.StatusCode = (int)status;
                response.ContentType = "application/json; charset=utf-8";
                response.ContentLength64 = bytes.Length;
                response.OutputStream.Write(bytes, 0, bytes.Length);
            }
            finally
            {
                response.Close();
            }
        }

        private static void WriteLog(string message)
        {
            try
            {
                var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelCalibrationAddin", "automation-bridge.log");
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                lock (LogLock)
                {
                    File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + message + Environment.NewLine);
                }
            }
            catch
            {
            }
        }
    }

    internal sealed class BridgeOptions
    {
        public int PublicPort { get; private set; } = 30771;
        public int InnerPort { get; private set; } = 30772;
        public int ExcelProcessId { get; private set; }
        public string Token { get; private set; } = string.Empty;

        public static BridgeOptions Parse(string[] args)
        {
            var options = new BridgeOptions();
            foreach (var argument in args ?? new string[0])
            {
                if (argument.StartsWith("--public-port=", StringComparison.OrdinalIgnoreCase))
                {
                    int.TryParse(argument.Substring(14), out var port);
                    if (port >= 1024 && port <= 65535) options.PublicPort = port;
                }
                else if (argument.StartsWith("--inner-port=", StringComparison.OrdinalIgnoreCase))
                {
                    int.TryParse(argument.Substring(13), out var port);
                    if (port >= 1024 && port <= 65535) options.InnerPort = port;
                }
                else if (argument.StartsWith("--excel-pid=", StringComparison.OrdinalIgnoreCase))
                {
                    int.TryParse(argument.Substring(12), out var processId);
                    options.ExcelProcessId = processId;
                }
                else if (argument.StartsWith("--token=", StringComparison.OrdinalIgnoreCase))
                {
                    options.Token = argument.Substring(8).Trim().Trim('"');
                }
            }

            if (options.InnerPort == options.PublicPort)
            {
                throw new ArgumentException("Invalid bridge options.");
            }

            return options;
        }
    }
}
