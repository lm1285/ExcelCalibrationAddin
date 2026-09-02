using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.ViewModels;
using ExcelCalibrationAddin.Host.Vsto;
using ExcelCalibrationAddin.Vsto.TaskPane;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private void EnsureFacade()
        {
            if (_facade != null)
            {
                return;
            }

            lock (_facadeInitializationLock)
            {
                if (_facade != null)
                {
                    return;
                }

                var loader = new ConfigurationLoader();
                var configuration = loader.Load(_configPath);
                configuration.Backend.AuthorizationToken = new CloudSessionStore().LoadToken();
                configuration.Generation = LoadGenerationConfiguration();
                _facade = new VstoAddinFacade(configuration);
                Trace.WriteLine($"[VSTO] Facade initialized. Config={_configPath}");
            }
        }

        private GenerationConfiguration LoadGenerationConfiguration()
        {
            if (_generationConfiguration == null)
            {
                _generationConfiguration = new GenerationConfigurationStore().Load(_configPath);
            }

            return _generationConfiguration;
        }

        private void RecalculateWorkbook(Excel.Workbook workbook)
        {
            try
            {
                if (workbook != null)
                {
                    workbook.ForceFullCalculation = true;
                }
            }
            catch
            {
            }

            try
            {
                if (workbook != null)
                {
                    foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                    {
                        worksheet.Calculate();
                    }
                }
                return;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Workbook calculate failed: {ex.Message}");
            }

            try
            {
                this.Application?.Calculate();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Application calculate failed: {ex.Message}");
            }
        }

        private static bool IsSameRange(CellRange left, CellRange right)
        {
            if (left == null || right == null)
            {
                return false;
            }

            return string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase) &&
                left.StartRow == right.StartRow &&
                left.EndRow == right.EndRow &&
                left.StartColumn == right.StartColumn &&
                left.EndColumn == right.EndColumn;
        }

    }
}
