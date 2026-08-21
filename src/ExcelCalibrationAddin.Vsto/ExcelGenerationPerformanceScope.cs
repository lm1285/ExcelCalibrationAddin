using System;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class ExcelGenerationPerformanceScope : IDisposable
    {
        private readonly Excel.Application _application;
        private readonly bool _screenUpdating;
        private readonly bool _enableEvents;
        private readonly bool _displayAlerts;
        private readonly Excel.XlCalculation _calculation;
        private bool _disposed;

        public ExcelGenerationPerformanceScope(Excel.Application application)
        {
            _application = application;
            if (_application == null)
            {
                return;
            }

            _screenUpdating = SafeGet(() => _application.ScreenUpdating, true);
            _enableEvents = SafeGet(() => _application.EnableEvents, true);
            _displayAlerts = SafeGet(() => _application.DisplayAlerts, true);
            _calculation = SafeGet(() => _application.Calculation, Excel.XlCalculation.xlCalculationAutomatic);
            CalculationWasAutomatic = _calculation == Excel.XlCalculation.xlCalculationAutomatic;

            SafeSet(() => _application.ScreenUpdating = false, "ScreenUpdating");
            SafeSet(() => _application.EnableEvents = false, "EnableEvents");
            SafeSet(() => _application.DisplayAlerts = false, "DisplayAlerts");
            SafeSet(() => _application.Calculation = Excel.XlCalculation.xlCalculationManual, "Calculation");
        }

        public bool CalculationWasAutomatic { get; }

        public void Dispose()
        {
            if (_disposed || _application == null)
            {
                return;
            }

            _disposed = true;
            SafeSet(() => _application.DisplayAlerts = _displayAlerts, "DisplayAlerts");
            SafeSet(() => _application.EnableEvents = _enableEvents, "EnableEvents");
            SafeSet(() => _application.ScreenUpdating = _screenUpdating, "ScreenUpdating");

            SafeSet(() => _application.Calculation = _calculation, "Calculation");
        }

        private static T SafeGet<T>(Func<T> getter, T fallback)
        {
            try
            {
                return getter();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Excel performance scope read failed: {ex.Message}");
                return fallback;
            }
        }

        private static void SafeSet(Action setter, string name)
        {
            try
            {
                setter();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Excel performance scope set {name} failed: {ex.Message}");
            }
        }
    }
}
