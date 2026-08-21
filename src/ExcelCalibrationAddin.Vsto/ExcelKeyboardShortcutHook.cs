using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class ExcelKeyboardShortcutHook : NativeWindow, IDisposable
    {
        private const int WmHotKey = 0x0312;
        private const int WmShortcutMessage = 0x8061;
        private const int WmKeyDown = 0x0100;
        private const int WmSysKeyDown = 0x0104;
        private const int WmKeyUp = 0x0101;
        private const int WmSysKeyUp = 0x0105;
        private const int WhGetMessage = 3;
        private const int WhCallWndProc = 4;
        private const int HotKeyId = 0xEC61;
        private const uint ModAlt = 0x0001;
        private const uint ModControl = 0x0002;
        private const uint ModShift = 0x0004;
        private const uint ModNoRepeat = 0x4000;
        private const uint GaRoot = 2;

        private readonly Action _onShortcut;
        private readonly HookProc _callWndProc;
        private readonly HookProc _getMessageProc;
        private IntPtr _excelWindowHandle;
        private uint _excelProcessId;
        private IntPtr _excelThreadHookHandle;
        private IntPtr _excelThreadGetMessageHookHandle;
        private Keys _shortcutKey;
        private uint _registeredModifiers;
        private bool _shortcutConfigured;
        private bool _shortcutRegistered;
        private bool _syntheticControlDown;
        private bool _syntheticAltDown;
        private bool _syntheticShiftDown;
        private bool _syntheticShortcutDown;
        private bool _disposed;

        public ExcelKeyboardShortcutHook(Action onShortcut)
        {
            _onShortcut = onShortcut ?? throw new ArgumentNullException(nameof(onShortcut));
            _callWndProc = ExcelThreadWindowProc;
            _getMessageProc = ExcelThreadGetMessageProc;
        }

        public void Attach(IntPtr excelWindowHandle)
        {
            if (_disposed || excelWindowHandle == IntPtr.Zero)
            {
                return;
            }

            _excelWindowHandle = excelWindowHandle;
            GetWindowThreadProcessId(_excelWindowHandle, out _excelProcessId);
            if (Handle == IntPtr.Zero)
            {
                CreateHandle(new CreateParams { Caption = "ExcelCalibrationAddin.ShortcutWindow" });
            }

            if (_excelThreadHookHandle == IntPtr.Zero)
            {
                var moduleHandle = GetModuleHandle(Process.GetCurrentProcess().MainModule?.ModuleName);
                var threadId = GetWindowThreadProcessId(_excelWindowHandle, out _);
                _excelThreadHookHandle = SetWindowsHookEx(
                    WhCallWndProc,
                    _callWndProc,
                    moduleHandle,
                    threadId);
                Trace.WriteLine(_excelThreadHookHandle == IntPtr.Zero
                    ? $"[VSTO] Excel thread keyboard hook failed: Win32Error={Marshal.GetLastWin32Error()}"
                    : $"[VSTO] Excel thread keyboard hook installed. ThreadId={threadId}");

                _excelThreadGetMessageHookHandle = SetWindowsHookEx(
                    WhGetMessage,
                    _getMessageProc,
                    moduleHandle,
                    threadId);
                Trace.WriteLine(_excelThreadGetMessageHookHandle == IntPtr.Zero
                    ? $"[VSTO] Excel thread message-queue hook failed: Win32Error={Marshal.GetLastWin32Error()}"
                    : $"[VSTO] Excel thread message-queue hook installed. ThreadId={threadId}");
            }
        }

        public void SetShortcutKey(string shortcutKey)
        {
            UnregisterShortcut();
            _shortcutKey = Keys.None;
            if (!TryParseShortcutKey(shortcutKey, out var key, out var modifiers) || Handle == IntPtr.Zero)
            {
                return;
            }

            _shortcutKey = key;
            _registeredModifiers = modifiers;
            _shortcutConfigured = true;
            _shortcutRegistered = RegisterHotKey(
                Handle,
                HotKeyId,
                modifiers | ModNoRepeat,
                (uint)key) != 0;
            if (_shortcutRegistered)
            {
                Trace.WriteLine($"[VSTO] Windows shortcut registered: {shortcutKey}");
            }
            else
            {
                Trace.WriteLine($"[VSTO] Windows shortcut registration failed: {shortcutKey}, Win32Error={Marshal.GetLastWin32Error()}");
            }
        }

        protected override void WndProc(ref Message message)
        {
            var isRegisteredHotKey = message.Msg == WmHotKey && message.WParam.ToInt32() == HotKeyId;
            var isExcelThreadShortcut = message.Msg == WmShortcutMessage;
            if (isRegisteredHotKey && IsExcelForeground() || isExcelThreadShortcut)
            {
                try
                {
                    _onShortcut();
                }
                catch (Exception exception)
                {
                    Trace.WriteLine($"[VSTO] Keyboard shortcut callback failed: {exception}");
                }

                message.Result = IntPtr.Zero;
                return;
            }

            base.WndProc(ref message);
        }

        private IntPtr ExcelThreadWindowProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && !_disposed && _shortcutConfigured)
            {
                var windowMessage = Marshal.PtrToStructure<CallWndProcMessage>(lParam);
                ProcessExcelKeyboardMessage(
                    windowMessage.WindowHandle,
                    windowMessage.Message,
                    windowMessage.WParam,
                    "CallWndProc");
            }

            return CallNextHookEx(_excelThreadHookHandle, code, wParam, lParam);
        }

        private IntPtr ExcelThreadGetMessageProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && !_disposed && _shortcutConfigured)
            {
                var windowMessage = Marshal.PtrToStructure<GetMessageData>(lParam);
                ProcessExcelKeyboardMessage(
                    windowMessage.WindowHandle,
                    windowMessage.Message,
                    windowMessage.WParam,
                    "GetMessage");
            }

            return CallNextHookEx(_excelThreadGetMessageHookHandle, code, wParam, lParam);
        }

        private void ProcessExcelKeyboardMessage(IntPtr windowHandle, int message, IntPtr wParam, string source)
        {
            TrackSyntheticModifier(message, wParam);
            var isKeyDown = message == WmKeyDown || message == WmSysKeyDown;
            var isKeyUp = message == WmKeyUp || message == WmSysKeyUp;
            var key = (Keys)(long)wParam;
            if (key != _shortcutKey || (!isKeyDown && !isKeyUp) || !IsExcelWindow(windowHandle))
            {
                return;
            }

            if (isKeyUp)
            {
                _syntheticShortcutDown = false;
                return;
            }

            var isSynthetic = IsSyntheticKey(key);
            var modifiersMatch = HasExpectedSyntheticModifiers();
            Trace.WriteLine(
                $"[VSTO] Excel target shortcut message: Source={source}, Key={key}, Synthetic={isSynthetic}, " +
                $"ModifiersMatch={modifiersMatch}, Foreground={IsExcelForeground()}");
            if (!_syntheticShortcutDown && isSynthetic && modifiersMatch)
            {
                _syntheticShortcutDown = true;
                PostMessage(Handle, WmShortcutMessage, IntPtr.Zero, IntPtr.Zero);
            }
        }

        private void TrackSyntheticModifier(int message, IntPtr wParam)
        {
            var isKeyDown = message == WmKeyDown || message == WmSysKeyDown;
            var isKeyUp = message == WmKeyUp || message == WmSysKeyUp;
            if (!isKeyDown && !isKeyUp)
            {
                return;
            }

            var key = (Keys)(long)wParam;
            var value = isKeyDown;
            if (key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey)
            {
                _syntheticControlDown = value;
            }
            else if (key == Keys.Menu || key == Keys.LMenu || key == Keys.RMenu)
            {
                _syntheticAltDown = value;
            }
            else if (key == Keys.ShiftKey || key == Keys.LShiftKey || key == Keys.RShiftKey)
            {
                _syntheticShiftDown = value;
            }
        }

        private bool HasExpectedSyntheticModifiers()
        {
            var modifiers = Keys.None;
            if (_syntheticControlDown) modifiers |= Keys.Control;
            if (_syntheticAltDown) modifiers |= Keys.Alt;
            if (_syntheticShiftDown) modifiers |= Keys.Shift;
            return modifiers == ToKeys(_registeredModifiers);
        }

        private static Keys ToKeys(uint modifiers)
        {
            var result = Keys.None;
            if ((modifiers & ModControl) != 0) result |= Keys.Control;
            if ((modifiers & ModAlt) != 0) result |= Keys.Alt;
            if ((modifiers & ModShift) != 0) result |= Keys.Shift;
            return result;
        }

        private static bool IsSyntheticKey(Keys key)
        {
            return (GetAsyncKeyState((int)key) & 0x8000) == 0;
        }

        private static bool TryParseShortcutKey(string shortcutKey, out Keys key, out uint modifiers)
        {
            key = Keys.None;
            modifiers = 0;
            if (string.IsNullOrWhiteSpace(shortcutKey)) return false;
            var parts = shortcutKey.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return false;

            for (var index = 0; index < parts.Length - 1; index++)
            {
                var modifier = parts[index].Trim();
                if (string.Equals(modifier, "Ctrl", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(modifier, "Control", StringComparison.OrdinalIgnoreCase)) modifiers |= ModControl;
                else if (string.Equals(modifier, "Alt", StringComparison.OrdinalIgnoreCase)) modifiers |= ModAlt;
                else if (string.Equals(modifier, "Shift", StringComparison.OrdinalIgnoreCase)) modifiers |= ModShift;
                else return false;
            }

            if (!Enum.TryParse(parts[parts.Length - 1].Trim(), true, out key) ||
                key == Keys.None || key == Keys.Control || key == Keys.Menu || key == Keys.ShiftKey ||
                (modifiers == 0 && (key < Keys.F1 || key > Keys.F12)))
            {
                key = Keys.None;
                modifiers = 0;
                return false;
            }

            return true;
        }

        private bool IsExcelForeground()
        {
            var foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return false;
            var rootWindow = GetAncestor(foreground, GaRoot);
            return IsExcelWindow(rootWindow);
        }

        private bool IsExcelWindow(IntPtr windowHandle)
        {
            if (windowHandle == IntPtr.Zero) return false;
            var rootWindow = GetAncestor(windowHandle, GaRoot);
            if (rootWindow == _excelWindowHandle) return true;
            uint processId;
            GetWindowThreadProcessId(rootWindow, out processId);
            if (processId != _excelProcessId) return false;
            var className = new StringBuilder(64);
            return GetClassName(rootWindow, className, className.Capacity) > 0 &&
                string.Equals(className.ToString(), "XLMAIN", StringComparison.OrdinalIgnoreCase);
        }

        private void UnregisterShortcut()
        {
            if (_shortcutRegistered && Handle != IntPtr.Zero)
            {
                UnregisterHotKey(Handle, HotKeyId);
            }
            _shortcutRegistered = false;
            _shortcutConfigured = false;
            _registeredModifiers = 0;
            _syntheticShortcutDown = false;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            UnregisterShortcut();
            if (_excelThreadHookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_excelThreadHookHandle);
                _excelThreadHookHandle = IntPtr.Zero;
            }
            if (_excelThreadGetMessageHookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_excelThreadGetMessageHookHandle);
                _excelThreadGetMessageHookHandle = IntPtr.Zero;
            }
            DestroyHandle();
        }

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct CallWndProcMessage
        {
            public IntPtr WindowHandle;
            public int Message;
            public IntPtr WParam;
            public IntPtr LParam;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct GetMessageData
        {
            public IntPtr WindowHandle;
            public int Message;
            public IntPtr WParam;
            public IntPtr LParam;
            public uint Time;
            public MessagePoint Point;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MessagePoint
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int RegisterHotKey(IntPtr windowHandle, int id, uint modifiers, uint virtualKeyCode);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnregisterHotKey(IntPtr windowHandle, int id);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int hookType, HookProc callback, IntPtr moduleHandle, uint threadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hookHandle);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hookHandle, int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr windowHandle, int message, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr windowHandle, uint flags);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr windowHandle, out uint processId);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int virtualKeyCode);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr windowHandle, StringBuilder className, int maxCount);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string moduleName);
    }
}
