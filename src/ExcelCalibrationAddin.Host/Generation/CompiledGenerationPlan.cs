using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Host.Generation
{
    public sealed class CompiledGenerationPlan
    {
        public string WorkbookKey { get; }
        public string ExactFingerprint { get; }
        public IReadOnlyList<MeasurementRule> Rules { get; }
        public DateTime CreatedUtc { get; }

        public CompiledGenerationPlan(string workbookKey, TemplateFingerprint fingerprint, IEnumerable<MeasurementRule> rules)
        {
            WorkbookKey = workbookKey ?? string.Empty;
            ExactFingerprint = fingerprint?.ExactFingerprint ?? string.Empty;
            Rules = (rules ?? Enumerable.Empty<MeasurementRule>()).Where(x => x != null).ToList().AsReadOnly();
            CreatedUtc = DateTime.UtcNow;
        }

        public bool IsValidFor(string workbookKey, TemplateFingerprint fingerprint)
        {
            return string.Equals(WorkbookKey, workbookKey ?? string.Empty, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(ExactFingerprint, fingerprint?.ExactFingerprint ?? string.Empty, StringComparison.OrdinalIgnoreCase) && Rules.Count > 0;
        }
    }
}
