using System;
using System.Collections.Generic;
using System.Net.Http;
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

        public TemplateSyncClient(HttpClient httpClient, string apiPrefix)
        {
            _httpClient = httpClient;
            _apiPrefix = apiPrefix.TrimEnd('/');
        }

        public async Task<string> MatchAsync(TemplateFingerprint fingerprint)
        {
            return await PostJsonAsync($"{_apiPrefix}/match", new { fingerprint });
        }

        public async Task<string> ListTemplatesAsync()
        {
            try
            {
                return await _httpClient.GetStringAsync($"{_apiPrefix}/list");
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
                using (var response = await _httpClient.PostAsync(url, new StringContent(json, Encoding.UTF8, "application/json")))
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
            catch (TaskCanceledException)
            {
                throw new InvalidOperationException("连接模板服务超时，请确认后端服务是否已启动。");
            }
            catch (HttpRequestException ex)
            {
                throw new InvalidOperationException("无法连接模板服务，请确认后端服务是否可访问。", ex);
            }
        }
    }
}
