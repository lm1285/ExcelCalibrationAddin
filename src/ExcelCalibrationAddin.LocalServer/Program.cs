using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Core.Services;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.LocalServer
{
    internal static class Program
    {
        private const int MaxRequestBodyBytes = 5 * 1024 * 1024;

        private static int Main(string[] args)
        {
            AddinFileLogger.Configure("LocalServer");
            try
            {
                var configurationPath = ResolveConfigurationPath(args);
                var configuration = new ConfigurationLoader().Load(configurationPath);
                var repository = new LocalTemplateRuleCacheRepository(configuration.Cache.SqliteFile);
                repository.Initialize();
                Trace.WriteLine($"[LocalServer] SQLite database initialized. Path={configuration.Cache.SqliteFile}");

                var prefix = BuildListenerPrefix(configuration.Backend.BaseUrl);
                using (var listener = new HttpListener())
                {
                    listener.Prefixes.Add(prefix);
                    listener.Start();
                    Console.WriteLine($"Excel template local server listening on {prefix}");
                    Console.WriteLine($"SQLite database: {configuration.Cache.SqliteFile}");
                    Console.WriteLine($"Log file: {AddinFileLogger.LogFilePath}");
                    Console.WriteLine("Press Ctrl+C to stop.");
                    Trace.WriteLine($"[LocalServer] Listening on {prefix}");

                    while (true)
                    {
                        var context = listener.GetContext();
                        HandleRequest(context, repository, configuration.Backend.TemplateApiPrefix);
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[LocalServer] Fatal error: {ex}");
                Console.Error.WriteLine(ex);
                return 1;
            }
        }

        private static string ResolveConfigurationPath(IReadOnlyList<string> args)
        {
            var configArgument = args.FirstOrDefault(item => item.StartsWith("--config=", StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(configArgument))
            {
                return configArgument.Substring("--config=".Length).Trim('"');
            }

            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
        }

        private static string BuildListenerPrefix(string baseUrl)
        {
            var uri = new Uri(string.IsNullOrWhiteSpace(baseUrl) ? "http://localhost:3002" : baseUrl);
            if (!string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                !uri.IsLoopback)
            {
                throw new InvalidOperationException("本地模板服务只允许监听 HTTP 回环地址。");
            }

            return new UriBuilder(Uri.UriSchemeHttp, uri.Host, uri.Port, "/").Uri.AbsoluteUri;
        }

        private static void HandleRequest(
            HttpListenerContext context,
            LocalTemplateRuleCacheRepository repository,
            string apiPrefix)
        {
            try
            {
                RejectBrowserOrigin(context.Request);
                var routeForLog = NormalizeRoute(context.Request.Url.AbsolutePath, apiPrefix);
                Trace.WriteLine($"[LocalServer] Request {context.Request.HttpMethod} {context.Request.Url.AbsolutePath} route={routeForLog}");

                var route = routeForLog;
                if (string.Equals(context.Request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase) && route == "/health")
                {
                    Trace.WriteLine("[LocalServer] Response 200 health");
                    WriteJson(context.Response, HttpStatusCode.OK, new { ok = true });
                    return;
                }

                if (string.Equals(context.Request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase) && route == "/list")
                {
                    var templates = repository.ListSavedTemplates();
                    Trace.WriteLine($"[LocalServer] Response 200 list count={templates.Count}");
                    WriteJson(context.Response, HttpStatusCode.OK, new { data = BuildTemplateCatalog(templates) });
                    return;
                }

                if (string.Equals(context.Request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase) && route == "/match")
                {
                    RequireJsonContentType(context.Request);
                    var request = ReadJson<MatchRequest>(context.Request);
                    if (request?.Fingerprint == null)
                    {
                        throw new RequestValidationException("Template fingerprint is required.");
                    }
                    var match = repository.FindBestMatch(request?.Fingerprint);
                    Trace.WriteLine($"[LocalServer] Response 200 match found={match != null} score={match?.MatchScore ?? 0:F0}");
                    WriteJson(context.Response, HttpStatusCode.OK, BuildMatchResponse(match));
                    return;
                }

                if (string.Equals(context.Request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase) && route == "/save")
                {
                    RequireJsonContentType(context.Request);
                    var request = ReadJson<SaveTemplateRequest>(context.Request);
                    ValidateSaveRequest(request);
                    var existing = repository.ListSavedTemplates().FirstOrDefault(item =>
                        (!string.IsNullOrWhiteSpace(request.TemplateId) &&
                         string.Equals(item.RemoteTemplateId, request.TemplateId, StringComparison.OrdinalIgnoreCase)) ||
                        string.Equals(item.TemplateName, request.TemplateName, StringComparison.OrdinalIgnoreCase));
                    var templateId = request.CreateNew || existing == null
                        ? $"template:{Guid.NewGuid():N}"
                        : existing.RemoteTemplateId;
                    var version = Math.Max(1, (existing?.RemoteVersion ?? 0) + 1);
                    repository.SaveAcceptedRemoteTemplate(
                        templateId,
                        request.TemplateName,
                        version,
                        DateTime.UtcNow,
                        null,
                        TemplateLifecycleStatus.Enabled,
                        request.Fingerprint,
                        request.Rules,
                        request.GenerationConfiguration,
                        request.DirectoryMetadata);
                    Trace.WriteLine($"[LocalServer] Response 200 save name={request.TemplateName} rules={request.Rules?.Count ?? 0}");
                    var saved = repository.FindByExactFingerprint(request.Fingerprint.ExactFingerprint);
                    var catalog = BuildTemplateCatalog(saved == null
                        ? new List<CachedTemplateRule>()
                        : new List<CachedTemplateRule> { saved });
                    WriteJson(context.Response, HttpStatusCode.OK, new { ok = true, data = catalog.FirstOrDefault() });
                    return;
                }

                Trace.WriteLine($"[LocalServer] Response 404 route={route}");
                WriteJson(context.Response, HttpStatusCode.NotFound, new { error = "Unknown endpoint." });
            }
            catch (RequestValidationException ex)
            {
                Trace.WriteLine($"[LocalServer] Request rejected: {ex.Message}");
                WriteJson(context.Response, HttpStatusCode.BadRequest, new { error = ex.Message });
            }
            catch (JsonException ex)
            {
                Trace.WriteLine($"[LocalServer] Invalid JSON: {ex.Message}");
                WriteJson(context.Response, HttpStatusCode.BadRequest, new { error = "Request body is not valid JSON." });
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[LocalServer] Request failed: {ex}");
                WriteJson(context.Response, HttpStatusCode.InternalServerError, new { error = "The local template service could not process the request." });
            }
        }

        private static string NormalizeRoute(string absolutePath, string apiPrefix)
        {
            var path = (absolutePath ?? string.Empty).TrimEnd('/');
            var prefix = (apiPrefix ?? string.Empty).TrimEnd('/');
            if (!string.IsNullOrWhiteSpace(prefix) &&
                path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                path = path.Substring(prefix.Length);
            }

            return string.IsNullOrWhiteSpace(path) ? "/" : path;
        }

        private static IReadOnlyList<object> BuildTemplateCatalog(IReadOnlyList<CachedTemplateRule> templates)
        {
            return (templates ?? new List<CachedTemplateRule>())
                .Select(template => new
                {
                    id = string.IsNullOrWhiteSpace(template.RemoteTemplateId) ? template.ExactFingerprint : template.RemoteTemplateId,
                    template_id = string.IsNullOrWhiteSpace(template.RemoteTemplateId) ? template.ExactFingerprint : template.RemoteTemplateId,
                    name = template.TemplateName,
                    template_name = template.TemplateName,
                    version = Math.Max(1, template.RemoteVersion),
                    status = (int)template.Status,
                    updated_at = (template.RemoteUpdatedAt ?? template.UpdatedAt).ToUniversalTime().ToString("o"),
                    deleted_at = template.DeletedAt?.ToUniversalTime().ToString("o") ?? string.Empty,
                    local_sync_status = (int)template.LocalSyncStatus,
                    sync_error = template.SyncError ?? string.Empty,
                    fingerprint_hash = JsonConvert.SerializeObject(template.Fingerprint),
                    directory_metadata = JsonConvert.SerializeObject(template.DirectoryMetadata),
                    rules_json = JsonConvert.SerializeObject(template.Rules ?? new List<MeasurementRule>()),
                    generation_config_json = template.GenerationConfiguration == null
                        ? string.Empty
                        : JsonConvert.SerializeObject(template.GenerationConfiguration),
                })
                .Cast<object>()
                .ToList();
        }

        private static object BuildMatchResponse(CachedTemplateRule match)
        {
            if (match == null)
            {
                return new { found = false, templates = new object[0] };
            }

            return new
            {
                found = true,
                templates = new[]
                {
                    new
                    {
                        id = string.IsNullOrWhiteSpace(match.RemoteTemplateId) ? match.ExactFingerprint : match.RemoteTemplateId,
                        template_id = string.IsNullOrWhiteSpace(match.RemoteTemplateId) ? match.ExactFingerprint : match.RemoteTemplateId,
                        name = match.TemplateName,
                        template_name = match.TemplateName,
                        version = Math.Max(1, match.RemoteVersion),
                        matchScore = match.MatchScore,
                        matchReason = match.MatchReason,
                        status = (int)match.Status,
                        updated_at = (match.RemoteUpdatedAt ?? match.UpdatedAt).ToUniversalTime().ToString("o"),
                        deleted_at = match.DeletedAt?.ToUniversalTime().ToString("o") ?? string.Empty,
                        fingerprint_hash = JsonConvert.SerializeObject(match.Fingerprint),
                        directory_metadata = JsonConvert.SerializeObject(match.DirectoryMetadata),
                        rules_json = JsonConvert.SerializeObject(match.Rules ?? new List<MeasurementRule>()),
                        generation_config_json = match.GenerationConfiguration == null
                            ? string.Empty
                            : JsonConvert.SerializeObject(match.GenerationConfiguration)
                    }
                }
            };
        }

        private static T ReadJson<T>(HttpListenerRequest request)
        {
            if (request.ContentLength64 > MaxRequestBodyBytes)
            {
                throw new RequestValidationException("请求内容超过模板同步允许的大小限制。");
            }

            using (var buffer = new MemoryStream())
            {
                var chunk = new byte[8192];
                var totalBytes = 0;
                int bytesRead;
                while ((bytesRead = request.InputStream.Read(chunk, 0, chunk.Length)) > 0)
                {
                    totalBytes += bytesRead;
                    if (totalBytes > MaxRequestBodyBytes)
                    {
                        throw new RequestValidationException("请求内容超过模板同步允许的大小限制。");
                    }

                    buffer.Write(chunk, 0, bytesRead);
                }

                var body = (request.ContentEncoding ?? Encoding.UTF8).GetString(buffer.ToArray());
                return string.IsNullOrWhiteSpace(body) ? default(T) : JsonConvert.DeserializeObject<T>(body);
            }
        }

        private static void RejectBrowserOrigin(HttpListenerRequest request)
        {
            if (!string.IsNullOrWhiteSpace(request.Headers["Origin"]))
            {
                throw new RequestValidationException("Browser-originated requests are not allowed.");
            }
        }

        private static void RequireJsonContentType(HttpListenerRequest request)
        {
            var contentType = request.ContentType ?? string.Empty;
            if (!contentType.StartsWith("application/json", StringComparison.OrdinalIgnoreCase))
            {
                throw new RequestValidationException("Content-Type must be application/json.");
            }
        }

        private static void ValidateSaveRequest(SaveTemplateRequest request)
        {
            if (request == null)
            {
                throw new RequestValidationException("Request body is required.");
            }

            if (string.IsNullOrWhiteSpace(request.TemplateName))
            {
                throw new RequestValidationException("Template name is required.");
            }

            if (string.IsNullOrWhiteSpace(request.Fingerprint?.ExactFingerprint))
            {
                throw new RequestValidationException("Template fingerprint is required.");
            }

            if (request.Rules == null || request.Rules.Count == 0)
            {
                throw new RequestValidationException("At least one template rule is required.");
            }
        }

        private static void WriteJson(HttpListenerResponse response, HttpStatusCode statusCode, object payload)
        {
            response.StatusCode = (int)statusCode;
            if (statusCode == HttpStatusCode.NoContent)
            {
                response.Close();
                return;
            }

            response.ContentType = "application/json; charset=utf-8";
            var bytes = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(payload ?? new { }));
            response.ContentLength64 = bytes.Length;
            response.OutputStream.Write(bytes, 0, bytes.Length);
            response.Close();
        }
    }

    internal sealed class RequestValidationException : Exception
    {
        public RequestValidationException(string message)
            : base(message)
        {
        }
    }

    internal sealed class MatchRequest
    {
        public TemplateFingerprint Fingerprint { get; set; }
    }

    internal sealed class SaveTemplateRequest
    {
        public string TemplateId { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public TemplateFingerprint Fingerprint { get; set; }
        public List<MeasurementRule> Rules { get; set; } = new List<MeasurementRule>();
        public GenerationConfiguration GenerationConfiguration { get; set; }
        public TemplateDirectoryMetadata DirectoryMetadata { get; set; }
        public bool CreateNew { get; set; }
    }
}
