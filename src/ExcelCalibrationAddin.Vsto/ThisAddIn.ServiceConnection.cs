using System;
using System.Net.Http;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Core.Services;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private async Task RefreshServiceConnectionStatusAsync()
        {
            var remoteConnected = await IsRemoteServiceAvailableAsync();
            var localConnected = await IsLocalAutomationServiceAvailableAsync();
            var remoteText = remoteConnected
                ? "\u8fdc\u7aef\u5df2\u8fde\u63a5"
                : "\u8fdc\u7aef\u672a\u8fde\u63a5";
            var localText = localConnected
                ? "\u672c\u673a\u5f71\u5200\u5df2\u5c31\u7eea"
                : "\u672c\u673a\u5f71\u5200\u672a\u5c31\u7eea";

            System.Diagnostics.Trace.WriteLine("[Yingdao] Service status: " + remoteText + " | " + localText);
        }

        private async Task<bool> IsRemoteServiceAvailableAsync()
        {
            try
            {
                var configuration = new ConfigurationLoader().Load(_configPath);
                if (string.IsNullOrWhiteSpace(configuration.Backend?.BaseUrl))
                {
                    return false;
                }

                var baseUri = new Uri(configuration.Backend.BaseUrl.TrimEnd('/') + "/");
                var healthUri = new Uri(baseUri, "api/health");
                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) })
                using (var response = await client.GetAsync(healthUri))
                {
                    return response.IsSuccessStatusCode;
                }
            }
            catch
            {
                return false;
            }
        }

        private async Task<bool> IsLocalAutomationServiceAvailableAsync()
        {
            try
            {
                var configuration = new ConfigurationLoader().Load(_configPath);
                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) })
                using (var response = await client.GetAsync(
                    "http://127.0.0.1:" + configuration.Automation.Port + "/api/yingdao/health"))
                {
                    return response.IsSuccessStatusCode;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
