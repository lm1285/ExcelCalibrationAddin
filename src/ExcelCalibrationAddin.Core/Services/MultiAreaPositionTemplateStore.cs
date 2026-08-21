using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelCalibrationAddin.Core.Models;
using Newtonsoft.Json;

namespace ExcelCalibrationAddin.Core.Services
{
    public sealed class MultiAreaPositionTemplateStore
    {
        private readonly object _syncRoot = new object();
        private readonly string _path;

        public MultiAreaPositionTemplateStore(string path = null)
        {
            _path = string.IsNullOrWhiteSpace(path) ? GetDefaultPath() : path;
        }

        public static string GetDefaultPath()
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppData, "ExcelCalibrationAddin", "multi-area-position-templates.json");
        }

        public IReadOnlyList<MultiAreaPositionTemplate> List()
        {
            lock (_syncRoot)
            {
                return LoadUnsafe()
                    .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                    .Select(item => item.Clone())
                    .ToList();
            }
        }

        public MultiAreaPositionTemplate Get(string id)
        {
            lock (_syncRoot)
            {
                var template = LoadUnsafe().FirstOrDefault(item =>
                    string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
                return template?.Clone();
            }
        }

        public MultiAreaPositionTemplate Save(string name, IEnumerable<AbsoluteAreaPosition> areas)
        {
            lock (_syncRoot)
            {
                var templates = LoadUnsafe();
                var existing = templates.FirstOrDefault(item =>
                    string.Equals(item.Name, (name ?? string.Empty).Trim(), StringComparison.CurrentCultureIgnoreCase));
                var saved = MultiAreaPositionTemplate.Create(
                    name,
                    areas,
                    existing?.Id,
                    existing?.CreatedUtc);

                if (existing != null)
                {
                    templates.Remove(existing);
                }

                templates.Add(saved);
                WriteUnsafe(templates);
                return saved.Clone();
            }
        }

        public MultiAreaPositionTemplate Save(MultiAreaPositionTemplate template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            template.Validate();
            lock (_syncRoot)
            {
                var templates = LoadUnsafe();
                var existing = templates.FirstOrDefault(item =>
                    string.Equals(item.Name, template.Name, StringComparison.CurrentCultureIgnoreCase));
                var saved = template.Clone();
                saved.Id = existing?.Id ?? (string.IsNullOrWhiteSpace(saved.Id) ? Guid.NewGuid().ToString("N") : saved.Id);
                saved.CreatedUtc = existing?.CreatedUtc ?? (saved.CreatedUtc == default(DateTime) ? DateTime.UtcNow : saved.CreatedUtc);
                saved.UpdatedUtc = DateTime.UtcNow;
                if (existing != null)
                {
                    templates.Remove(existing);
                }

                templates.Add(saved);
                WriteUnsafe(templates);
                return saved.Clone();
            }
        }

        public bool Delete(string id)
        {
            lock (_syncRoot)
            {
                var templates = LoadUnsafe();
                var removed = templates.RemoveAll(item =>
                    string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase)) > 0;
                if (removed)
                {
                    WriteUnsafe(templates);
                }

                return removed;
            }
        }

        private List<MultiAreaPositionTemplate> LoadUnsafe()
        {
            if (!File.Exists(_path))
            {
                return new List<MultiAreaPositionTemplate>();
            }

            var content = File.ReadAllText(_path);
            var templates = JsonConvert.DeserializeObject<List<MultiAreaPositionTemplate>>(content) ??
                            new List<MultiAreaPositionTemplate>();
            foreach (var template in templates)
            {
                template.Validate();
            }

            return templates
                .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.OrderByDescending(item => item.UpdatedUtc).First())
                .ToList();
        }

        private void WriteUnsafe(IReadOnlyCollection<MultiAreaPositionTemplate> templates)
        {
            var directory = Path.GetDirectoryName(_path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(_path, JsonConvert.SerializeObject(templates, Formatting.Indented));
        }
    }
}
