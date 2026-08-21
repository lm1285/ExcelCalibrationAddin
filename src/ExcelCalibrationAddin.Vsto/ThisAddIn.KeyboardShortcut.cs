using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private ExcelKeyboardShortcutHook _keyboardShortcutHook;
        private bool _shortcutGenerationRunning;

        private void InitializeKeyboardShortcut()
        {
            if (_keyboardShortcutHook == null)
            {
                _keyboardShortcutHook = new ExcelKeyboardShortcutHook(TriggerShortcutGeneration);
            }

            try
            {
                var configuration = LoadGenerationConfiguration();
                _keyboardShortcutHook.Attach(new IntPtr(Application.Hwnd));
                _keyboardShortcutHook.SetShortcutKey(configuration.GenerateShortcutKey);
                Trace.WriteLine($"[VSTO] Generation shortcut initialized: {configuration.GenerateShortcutKey}");
            }
            catch (Exception exception)
            {
                Trace.WriteLine($"[VSTO] Generation shortcut initialization failed: {exception}");
            }
        }

        private void ApplyKeyboardShortcutConfiguration()
        {
            try
            {
                InitializeKeyboardShortcut();
            }
            catch (Exception exception)
            {
                Trace.WriteLine($"[VSTO] Generation shortcut refresh failed: {exception}");
            }
        }

        private void TriggerShortcutGeneration()
        {
            if (_shortcutGenerationRunning)
            {
                return;
            }

            _shortcutGenerationRunning = true;
            _ = GenerateFromShortcutAsync();
        }

        private async Task GenerateFromShortcutAsync()
        {
            try
            {
                Trace.WriteLine("[VSTO] Keyboard shortcut: GenerateRandom");
                await GenerateRandomNumbersCurrentWorkbookAsync();
            }
            catch (Exception exception)
            {
                Trace.WriteLine($"[VSTO] Keyboard shortcut generate random failed: {exception}");
                MessageBox.Show(exception.Message, "生成随机数失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                _shortcutGenerationRunning = false;
            }
        }

        private void DisposeKeyboardShortcut()
        {
            _keyboardShortcutHook?.Dispose();
            _keyboardShortcutHook = null;
        }
    }
}
