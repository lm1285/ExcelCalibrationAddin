using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelCalibrationAddin.Host.Services
{
    public static class AutomaticMatchSheetSelector
    {
        public static IReadOnlyList<string> Select(
            IEnumerable<string> workbookSheetNames,
            IEnumerable<string> templateSheetNames,
            bool hasEnabledTemplates,
            bool hasTemplateWithoutSheetMetadata)
        {
            var workbookNames = Normalize(workbookSheetNames);
            if (!hasEnabledTemplates || workbookNames.Count == 0)
            {
                return new List<string>();
            }

            var templateNames = new HashSet<string>(
                Normalize(templateSheetNames),
                StringComparer.OrdinalIgnoreCase);
            var candidates = workbookNames
                .Where(templateNames.Contains)
                .ToList();

            if (hasTemplateWithoutSheetMetadata)
            {
                candidates.AddRange(workbookNames.Where(name =>
                    !candidates.Contains(name, StringComparer.OrdinalIgnoreCase)));
            }

            return candidates;
        }

        private static List<string> Normalize(IEnumerable<string> names)
        {
            return (names ?? Enumerable.Empty<string>())
                .Select(name => (name ?? string.Empty).Trim())
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }
}
