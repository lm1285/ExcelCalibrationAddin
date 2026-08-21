using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;

namespace ExcelCalibrationAddin.Host.UseCases
{
    public interface IWorkbookSnapshotProvider
    {
        WorkbookSnapshot Capture();
        WorkbookSnapshot Capture(IEnumerable<CellRange> ranges);
        string GetActiveSheetName();
    }

    public interface IFingerprintSnapshotProvider
    {
        WorkbookSnapshot CaptureFingerprint();
    }

    public sealed class TemplateRecognitionUseCase
    {
        private readonly IWorkbookSnapshotProvider _snapshotProvider;
        private readonly TemplateFingerprintBuilder _fingerprintBuilder;
        private readonly FieldMatcher _fieldMatcher;

        public TemplateRecognitionUseCase(
            IWorkbookSnapshotProvider snapshotProvider,
            TemplateFingerprintBuilder fingerprintBuilder,
            FieldMatcher fieldMatcher)
        {
            _snapshotProvider = snapshotProvider;
            _fingerprintBuilder = fingerprintBuilder;
            _fieldMatcher = fieldMatcher;
        }

        public RecognitionResult Execute()
        {
            var workbook = _snapshotProvider.Capture();
            RecognitionProgress.Report(60, "正在生成模板指纹...");
            var fingerprint = _fingerprintBuilder.Build(workbook);
            var fields = new List<RecognizedField>();
            var activeSheetName = _snapshotProvider.GetActiveSheetName();

            foreach (var sheet in workbook.Sheets)
            {
                if (!string.IsNullOrWhiteSpace(activeSheetName) &&
                    !string.Equals(sheet.Name, activeSheetName, System.StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                RecognitionProgress.Report(62, $"正在识别项目区域...");
                fields.AddRange(_fieldMatcher.MatchMeasurementFields(sheet));
            }

            return new RecognitionResult
            {
                Snapshot = workbook,
                Fingerprint = fingerprint,
                RecognizedFields = fields.OrderByDescending(item => item.Score).ToList()
            };
        }

        public RecognitionResult ExecuteFingerprintOnly()
        {
            var fingerprintSnapshotProvider = _snapshotProvider as IFingerprintSnapshotProvider;
            var workbook = fingerprintSnapshotProvider != null
                ? fingerprintSnapshotProvider.CaptureFingerprint()
                : _snapshotProvider.Capture();
            RecognitionProgress.Report(60, "正在生成模板指纹...");
            var fingerprint = _fingerprintBuilder.Build(workbook);

            return new RecognitionResult
            {
                Snapshot = workbook,
                Fingerprint = fingerprint,
                RecognizedFields = new List<RecognizedField>()
            };
        }
    }

    public sealed class RecognitionResult
    {
        public WorkbookSnapshot Snapshot { get; set; }
        public TemplateFingerprint Fingerprint { get; set; }
        public List<RecognizedField> RecognizedFields { get; set; } = new List<RecognizedField>();
    }
}
