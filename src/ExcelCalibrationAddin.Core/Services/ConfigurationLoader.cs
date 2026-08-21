using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class ConfigurationLoader
    {
        public PluginConfiguration Load(string configFilePath)
        {
            var configurationPath = FindConfigurationPath(configFilePath);
            if (string.IsNullOrWhiteSpace(configurationPath))
            {
                return Normalize(new PluginConfiguration());
            }

            var content = File.ReadAllText(configurationPath);
            var configuration = JsonConvert.DeserializeObject<PluginConfiguration>(content) ?? new PluginConfiguration();
            return Normalize(configuration);
        }

        public string ExpandPath(string path)
        {
            return Environment.ExpandEnvironmentVariables(path ?? string.Empty);
        }

        private static string FindConfigurationPath(string configFilePath)
        {
            var candidates = new List<string>();

            if (!string.IsNullOrWhiteSpace(configFilePath))
            {
                candidates.Add(configFilePath);
            }

            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (!string.IsNullOrWhiteSpace(baseDirectory))
            {
                candidates.Add(Path.Combine(baseDirectory, "appsettings.json"));
                candidates.Add(Path.Combine(baseDirectory, "appsettings.example.json"));
            }

            var assemblyDirectory = Path.GetDirectoryName(typeof(ConfigurationLoader).Assembly.Location);
            if (!string.IsNullOrWhiteSpace(assemblyDirectory))
            {
                candidates.Add(Path.Combine(assemblyDirectory, "appsettings.json"));
                candidates.Add(Path.Combine(assemblyDirectory, "appsettings.example.json"));
            }

            candidates.Add(Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json"));
            candidates.Add(Path.Combine(Directory.GetCurrentDirectory(), "appsettings.example.json"));

            return candidates.FirstOrDefault(File.Exists);
        }

        private PluginConfiguration Normalize(PluginConfiguration configuration)
        {
            if (configuration == null)
            {
                configuration = new PluginConfiguration();
            }

            if (configuration.Backend == null)
            {
                configuration.Backend = new BackendConfiguration();
            }

            if (configuration.Cache == null)
            {
                configuration.Cache = new CacheConfiguration();
            }

            if (configuration.Generation == null)
            {
                configuration.Generation = new GenerationConfiguration();
            }

            var backendDefaults = new BackendConfiguration();
            if (string.IsNullOrWhiteSpace(configuration.Backend.BaseUrl))
            {
                configuration.Backend.BaseUrl = backendDefaults.BaseUrl;
            }
            else
            {
                configuration.Backend.BaseUrl = configuration.Backend.BaseUrl.Trim();
            }

            if (string.IsNullOrWhiteSpace(configuration.Backend.TemplateApiPrefix))
            {
                configuration.Backend.TemplateApiPrefix = backendDefaults.TemplateApiPrefix;
            }
            else
            {
                configuration.Backend.TemplateApiPrefix = "/" + configuration.Backend.TemplateApiPrefix.Trim().Trim('/');
            }

            if (string.IsNullOrWhiteSpace(configuration.Cache.SqliteFile))
            {
                configuration.Cache.SqliteFile = new CacheConfiguration().SqliteFile;
            }

            configuration.Cache.SqliteFile = ExpandPath(configuration.Cache.SqliteFile);
            configuration.Generation = new GenerationConfigurationStore(this).Normalize(configuration.Generation);
            return configuration;
        }
    }
}
