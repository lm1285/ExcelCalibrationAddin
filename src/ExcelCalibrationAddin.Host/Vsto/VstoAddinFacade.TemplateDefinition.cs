using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.Generation;
using ExcelCalibrationAddin.Host.Interop;
using ExcelCalibrationAddin.Host.Recognition;
using ExcelCalibrationAddin.Host.Services;

namespace ExcelCalibrationAddin.Host.Vsto
{
    public sealed partial class VstoAddinFacade
    {
        public IReadOnlyList<MeasurementRule> PrepareRulesForTemplateSave(
            dynamic workbook,
            IReadOnlyList<MeasurementRule> rules)
        {
            var prepared = (rules ?? new List<MeasurementRule>())
                .Where(rule => rule != null)
                .Select(MeasurementRuleCloner.Clone)
                .ToList();
            if (workbook == null || prepared.Count == 0)
            {
                return prepared;
            }

            var snapshot = new ExcelInteropSnapshotProvider(workbook).Capture();
            var builder = new TemplateFieldDefinitionBuilder(new NumberFormatInterpreter());
            var structureAnalyzer = new MeasurementRuleStructureAnalyzer();
            var rowMappingBuilder = new RowMappingBuilder();
            foreach (var rule in prepared)
            {
                var sheetName = rule.TargetRange?.SheetName;
                var sheet = snapshot.Sheets.FirstOrDefault(item =>
                    string.Equals(item.Name, sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet == null) continue;
                var writableResolution = WritableCellResolver.Resolve(snapshot, rule.TargetRange);
                rule.WritableCells = writableResolution.Cells;
                rule.GroupSize = writableResolution.Cells.Count;
                rule.ErrorFormula = null;
                structureAnalyzer.Apply(snapshot, new[] { rule });
                rowMappingBuilder.Apply(snapshot, new[] { rule });
                rule.TemplateDefinition = builder.Build(sheet, rule);
            }

            return prepared;
        }
    }
}
