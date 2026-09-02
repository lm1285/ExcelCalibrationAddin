using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class YingdaoAutomationServer : IDisposable
    {
        private const string ApiPrefix = "/api/yingdao";
        private const int MaxRequestBodyBytes = 1024 * 1024;
        private readonly HttpListener _listener = new HttpListener();
        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();
        private readonly Func<Task<object>> _generate;
        private readonly Func<Task<object>> _status;
        private readonly string _token;
        private readonly SemaphoreSlim _generationGate = new SemaphoreSlim(1, 1);
        private Task _listenTask = Task.CompletedTask;

        public YingdaoAutomationServer(
            int port,
            string token,
            Func<Task<object>> generate,
            Func<Task<object>> status)
        {
            if (port < 1024 || port > 65535)
            {
                throw new ArgumentOutOfRangeException(nameof(port));
            }

            _token = token ?? string.Empty;
            _generate = generate ?? throw new ArgumentNullException(nameof(generate));
            _status = status ?? throw new ArgumentNullException(nameof(status));
            Port = port;
        }

        public int Port { get; }

        public bool IsListening => _listener.IsListening;

        public void Start()
        {
            _listener.Prefixes.Add($"http://127.0.0.1:{Port}/");
            _listener.Start();
            _listenTask = ListenAsync();
        }

        public void Dispose()
        {
            _cancellation.Cancel();
            if (_listener.IsListening)
            {
                _listener.Stop();
            }

            _listener.Close();
            _generationGate.Dispose();
            _cancellation.Dispose();
        }

        private async Task ListenAsync()
        {
            while (!_cancellation.IsCancellationRequested)
            {
                HttpListenerContext context;
                try
                {
                    context = await _listener.GetContextAsync().ConfigureAwait(false);
                }
                catch (HttpListenerException)
                {
                    break;
                }
                catch (ObjectDisposedException)
                {
                    break;
                }

                _ = HandleRequestAsync(context);
            }
        }

        private async Task HandleRequestAsync(HttpListenerContext context)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(context.Request.Headers["Origin"]))
                {
                    WriteJson(context.Response, HttpStatusCode.Forbidden, new { ok = false, error = "Browser requests are not allowed." });
                    return;
                }

                if (!IsAuthorized(context.Request))
                {
                    WriteJson(context.Response, HttpStatusCode.Unauthorized, new { ok = false, error = "Invalid automation token." });
                    return;
                }

                var route = NormalizeRoute(context.Request.Url?.AbsolutePath);
                if (context.Request.HttpMethod.Equals("GET", StringComparison.OrdinalIgnoreCase) && route == "/health")
                {
                    WriteJson(context.Response, HttpStatusCode.OK, new
                    {
                        ok = true,
                        service = "ExcelCalibrationAddin.Yingdao",
                        port = Port
                    });
                    return;
                }

                if (context.Request.HttpMethod.Equals("GET", StringComparison.OrdinalIgnoreCase) && route == "/status")
                {
                    WriteJson(context.Response, HttpStatusCode.OK, await _status().ConfigureAwait(false));
                    return;
                }

                if (context.Request.HttpMethod.Equals("POST", StringComparison.OrdinalIgnoreCase) && route == "/generate")
                {
                    RequireJsonContentType(context.Request);
                    ReadAndValidateBody(context.Request);
                    await _generationGate.WaitAsync().ConfigureAwait(false);
                    try
                    {
                        WriteJson(context.Response, HttpStatusCode.OK, await _generate().ConfigureAwait(false));
                    }
                    finally
                    {
                        _generationGate.Release();
                    }

                    return;
                }

                WriteJson(context.Response, HttpStatusCode.NotFound, new { ok = false, error = "Unknown endpoint." });
            }
            catch (InvalidOperationException ex)
            {
                WriteJson(context.Response, HttpStatusCode.Conflict, new { ok = false, error = ex.Message });
            }
            catch (ArgumentException ex)
            {
                WriteJson(context.Response, HttpStatusCode.BadRequest, new { ok = false, error = ex.Message });
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Yingdao] Request failed: {ex}");
                WriteJson(context.Response, HttpStatusCode.InternalServerError, new { ok = false, error = ex.Message });
            }
        }

        private bool IsAuthorized(HttpListenerRequest request)
        {
            return string.IsNullOrWhiteSpace(_token) ||
                string.Equals(request.Headers["X-Excel-Calibration-Token"], _token, StringComparison.Ordinal);
        }

        private static string NormalizeRoute(string absolutePath)
        {
            var path = (absolutePath ?? string.Empty).TrimEnd('/');
            if (path.StartsWith(ApiPrefix, StringComparison.OrdinalIgnoreCase))
            {
                path = path.Substring(ApiPrefix.Length);
            }

            return string.IsNullOrWhiteSpace(path) ? "/" : path;
        }

        private static void RequireJsonContentType(HttpListenerRequest request)
        {
            if (!(request.ContentType ?? string.Empty).StartsWith("application/json", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Content-Type must be application/json.");
            }
        }

        private static void ReadAndValidateBody(HttpListenerRequest request)
        {
            if (request.ContentLength64 > MaxRequestBodyBytes)
            {
                throw new ArgumentException("Request body is too large.");
            }

            var buffer = new byte[8192];
            var total = 0;
            int read;
            while ((read = request.InputStream.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += read;
                if (total > MaxRequestBodyBytes)
                {
                    throw new ArgumentException("Request body is too large.");
                }
            }
        }

        private static void WriteJson(HttpListenerResponse response, HttpStatusCode status, object payload)
        {
            var serializer = new JavaScriptSerializer();
            var normalizedPayload = payload is YingdaoAutomationStatus automationStatus
                ? new
                {
                    ok = automationStatus.Ok,
                    addinLoaded = automationStatus.AddinLoaded,
                    excelRunning = automationStatus.ExcelRunning,
                    workbookOpen = automationStatus.WorkbookOpen,
                    workbookName = automationStatus.WorkbookName,
                    templateMatched = automationStatus.TemplateMatched,
                    canGenerate = automationStatus.CanGenerate,
                    templateName = automationStatus.TemplateName,
                    exactFingerprint = automationStatus.ExactFingerprint,
                    ruleCount = automationStatus.RuleCount,
                    message = automationStatus.Message
                }
                : payload ?? new Dictionary<string, object>();
            var bytes = Encoding.UTF8.GetBytes(serializer.Serialize(normalizedPayload));
            response.StatusCode = (int)status;
            response.ContentType = "application/json; charset=utf-8";
            response.ContentEncoding = Encoding.UTF8;
            response.ContentLength64 = bytes.Length;
            try
            {
                response.OutputStream.Write(bytes, 0, bytes.Length);
            }
            finally
            {
                response.Close();
            }
        }
    }

    internal sealed class YingdaoAutomationStatus
    {
        public bool Ok { get; set; }
        public bool AddinLoaded { get; set; }
        public bool ExcelRunning { get; set; }
        public bool WorkbookOpen { get; set; }
        public string WorkbookName { get; set; } = string.Empty;
        public bool TemplateMatched { get; set; }
        public bool CanGenerate { get; set; }
        public string TemplateName { get; set; } = string.Empty;
        public string ExactFingerprint { get; set; } = string.Empty;
        public int RuleCount { get; set; }
        public string Message { get; set; } = string.Empty;
    }
}
