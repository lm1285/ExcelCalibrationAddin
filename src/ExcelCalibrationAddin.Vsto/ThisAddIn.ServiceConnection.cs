using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using ExcelCalibrationAddin.Core.Services;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        internal async Task RefreshServiceConnectionStatusAsync()
        {
            var configuration = new ConfigurationLoader().Load(_configPath);
            var address = configuration.Backend?.BaseUrl?.TrimEnd('/') ?? string.Empty;
            var ip = ResolveCloudIp(address);
            var health = await QueryRemoteHealthAsync(address);
            var cloudText = health.Connected ? "已连接" : "未连接";
            var databaseText = health.DatabaseConnected ? "已连接" : "未连接";
            var loginText = health.Authenticated ? "已登录" : "未登录";

            ApplyServiceStatus(cloudText, address, ip, databaseText, loginText);
            System.Diagnostics.Trace.WriteLine(
                "[Cloud] Service status: cloud=" + cloudText + ", database=" + databaseText + ", login=" + loginText + ", address=" + address + ", ip=" + ip);
        }

        private async Task<RemoteHealthStatus> QueryRemoteHealthAsync(string address)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(address))
                {
                    return new RemoteHealthStatus();
                }

                var baseUri = new Uri(address.TrimEnd('/') + "/");
                var healthUri = new Uri(baseUri, "api/health");
                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) })
                using (var response = await client.GetAsync(healthUri))
                {
                    var body = await response.Content.ReadAsStringAsync();
                    var payload = new JavaScriptSerializer().DeserializeObject(body) as Dictionary<string, object>;
                    var dbStatus = payload != null && payload.ContainsKey("dbStatus")
                        ? Convert.ToString(payload["dbStatus"])
                        : string.Empty;
                    return new RemoteHealthStatus
                    {
                        Connected = response.IsSuccessStatusCode,
                        DatabaseConnected = string.Equals(dbStatus, "connected", StringComparison.OrdinalIgnoreCase),
                        Authenticated = await IsCloudAuthenticatedAsync(baseUri, new CloudSessionStore().LoadToken())
                    };
                }
            }
            catch
            {
                return new RemoteHealthStatus();
            }
        }

        private static string ResolveCloudIp(string address)
        {
            try
            {
                var host = new Uri(address).DnsSafeHost;
                var addresses = Dns.GetHostAddresses(host)
                    .Where(item => item.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    .Select(item => item.ToString())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                return addresses.Length == 0 ? "未知" : string.Join(", ", addresses);
            }
            catch
            {
                return "未知";
            }
        }

        private async Task<bool> IsCloudAuthenticatedAsync(Uri baseUri, string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, new Uri(baseUri, "api/auth/user")))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Trim());
                    using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) })
                    using (var response = await client.SendAsync(request))
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

        private void ApplyServiceStatus(string cloudStatus, string address, string ip, string databaseStatus, string loginStatus)
        {
            Action update = () => Globals.Ribbons?.CalibrationRibbon?.UpdateServiceStatus(
                cloudStatus,
                address,
                ip,
                databaseStatus,
                loginStatus);

            if (_excelUiSynchronizationContext == null || SynchronizationContext.Current == _excelUiSynchronizationContext)
            {
                update();
                return;
            }

            _excelUiSynchronizationContext.Post(_ => update(), null);
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

    internal sealed class RemoteHealthStatus
    {
        public bool Connected { get; set; }
        public bool DatabaseConnected { get; set; }
        public bool Authenticated { get; set; }
    }
}
