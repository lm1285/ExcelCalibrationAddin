using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class TemplateSyncClient
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiPrefix;

        public TemplateSyncClient(HttpClient httpClient, string apiPrefix, string authorizationToken = null)
        {
            _httpClient = httpClient;
            _apiPrefix = apiPrefix.TrimEnd('/');
            AuthorizationToken = authorizationToken ?? string.Empty;
        }

        public string AuthorizationToken { get; set; }

        public async Task<string> MatchAsync(TemplateFingerprint fingerprint)
        {
            return await PostJsonAsync($"{_apiPrefix}/match", new { fingerprint });
        }

        public async Task<string> ListTemplatesAsync()
        {
            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, $"{_apiPrefix}/list"))
                {
                    AddAuthorization(request);
                    using (var response = await _httpClient.SendAsync(request))
                    {
                        var body = await response.Content.ReadAsStringAsync();
                        if (!response.IsSuccessStatusCode)
                        {
                            throw new InvalidOperationException($"模板服务返回错误：{body}");
                        }

                        return body;
                    }
                }
            }
            catch (TaskCanceledException)
            {
                throw new InvalidOperationException("连接模板服务超时，请确认后端服务是否已启动。");
            }
            catch (HttpRequestException ex)
            {
                throw new InvalidOperationException("无法连接模板服务，请确认后端服务是否可访问。", ex);
            }
        }

        public async Task<string> SaveTemplateAsync(
            string templateName,
            TemplateFingerprint fingerprint,
            IReadOnlyList<MeasurementRule> rules,
            GenerationConfiguration generationConfiguration,
            string templateId = null,
            bool createNew = false,
            TemplateDirectoryMetadata directoryMetadata = null)
        {
            return await PostJsonAsync($"{_apiPrefix}/save", new
            {
                templateId,
                templateName,
                fingerprint,
                rules,
                generationConfiguration,
                createNew,
                directoryMetadata
            });
        }

        private async Task<string> PostJsonAsync(string url, object payload)
        {
            var json = JsonConvert.SerializeObject(payload);

            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                    AddAuthorization(request);
                    using (var response = await _httpClient.SendAsync(request))
                    {
                        var responseBody = await response.Content.ReadAsStringAsync();
                        if (!response.IsSuccessStatusCode)
                        {
                            var message = string.IsNullOrWhiteSpace(responseBody)
                                ? response.ReasonPhrase
                                : responseBody;
                            throw new InvalidOperationException($"模板服务返回错误：{message}");
                        }

                        return responseBody;
                    }
                }
            }
            catch (TaskCanceledException)
            {
                throw new InvalidOperationException("连接模板服务超时，请确认后端服务是否已启动。");
            }
            catch (HttpRequestException ex)
            {
                throw new InvalidOperationException("无法连接模板服务，请确认后端服务是否可访问。", ex);
            }
        }

        private void AddAuthorization(HttpRequestMessage request)
        {
            var token = (AuthorizationToken ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(token))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            }
        }
    }
}
