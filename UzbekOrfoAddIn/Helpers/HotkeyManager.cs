using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Manages global keyboard shortcuts for the UzbekOrfo add-in using
    /// RegisterHotKey + a hidden message window.
    ///
    /// Hotkeys are processed only when a Word window is in the foreground.
    /// Actions are executed on the UI thread because the hidden window is
    /// created on that thread.
    ///
    /// Usage:
    ///   HotkeyManager.Register(new[] { new HotkeyDef(...), ... });
    ///   HotkeyManager.Unregister();
    /// </summary>
    public static class HotkeyManager
    {
        // =====================================================================
        //  WIN32
        // =====================================================================

        private const int WM_HOTKEY = 0x0312;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_NOREPEAT = 0x4000;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // =====================================================================
        //  PUBLIC TYPES
        // =====================================================================

        [Flags]
        public enum Modifiers : byte
        {
            None  = 0,
            Ctrl  = 1,
            Shift = 2,
            Alt   = 4
        }

        public sealed class HotkeyDef
        {
            public Modifiers Mod    { get; }
            public Keys      Key    { get; }
            public Action    Action { get; }
            public string    Label  { get; }

            public HotkeyDef(Modifiers mod, Keys key, Action action, string label = "")
            {
                Mod    = mod;
                Key    = key;
                Action = action;
                Label  = label;
            }
        }

        // =====================================================================
        //  STATE
        // =====================================================================

        private static readonly Dictionary<int, HotkeyDef> _registeredHotkeys = new Dictionary<int, HotkeyDef>();
        private static MessageWindow _window;
        private static int _nextId = 1;
        private static uint _wordProcId;

        // =====================================================================
        //  PUBLIC API
        // =====================================================================

        /// <summary>
        /// Registers all hotkeys. Must be called on the UI thread.
        /// </summary>
        public static void Register(IEnumerable<HotkeyDef> hotkeys)
        {
            Unregister();

            _wordProcId = (uint)Process.GetCurrentProcess().Id;

            try
            {
                _window = new MessageWindow();
                _window.HotkeyPressed += OnHotkeyPressed;

                int registeredCount = 0;
                foreach (var hotkey in hotkeys)
                {
                    int id = _nextId++;
                    uint modifiers = ToNativeModifiers(hotkey.Mod);
                    uint key = (uint)hotkey.Key;

                    if (RegisterHotKey(_window.Handle, id, modifiers, key))
                    {
                        _registeredHotkeys[id] = hotkey;
                        registeredCount++;
                    }
                    else
                    {
                        int err = Marshal.GetLastWin32Error();
                        Logger.Warn($"HotkeyManager: failed to register [{hotkey.Label}] (Win32 error {err}).");
                    }
                }

                Logger.Info($"HotkeyManager: registered {registeredCount} hotkeys.");
            }
            catch (Exception ex)
            {
                Logger.Error("HotkeyManager.Register failed", ex);
            }
        }

        public static void Unregister()
        {
            if (_window != null)
            {
                try
                {
                    foreach (var id in new List<int>(_registeredHotkeys.Keys))
                    {
                        try { UnregisterHotKey(_window.Handle, id); } catch { }
                    }
                }
                catch { }
            }

            _registeredHotkeys.Clear();

            if (_window != null)
            {
                try { _window.HotkeyPressed -= OnHotkeyPressed; } catch { }
                try { _window.Dispose(); } catch { }
                _window = null;
            }
        }

        // =====================================================================
        //  HOTKEY CALLBACK
        // =====================================================================

        private static void OnHotkeyPressed(int id)
        {
            try
            {
                if (!IsForegroundWordWindow())
                    return;

                HotkeyDef hotkey;
                if (!_registeredHotkeys.TryGetValue(id, out hotkey) || hotkey.Action == null)
                    return;

                hotkey.Action();
            }
            catch (Exception ex)
            {
                Logger.Error($"HotkeyManager: execution failed for id={id}", ex);
            }
        }

        // =====================================================================
        //  HELPERS
        // =====================================================================

        private static uint ToNativeModifiers(Modifiers modifiers)
        {
            uint native = MOD_NOREPEAT;

            if ((modifiers & Modifiers.Alt) != 0)
                native |= MOD_ALT;
            if ((modifiers & Modifiers.Ctrl) != 0)
                native |= MOD_CONTROL;
            if ((modifiers & Modifiers.Shift) != 0)
                native |= MOD_SHIFT;

            return native;
        }

        private sealed class MessageWindow : NativeWindow, IDisposable
        {
            public event Action<int> HotkeyPressed;

            public MessageWindow()
            {
                CreateHandle(new CreateParams());
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_HOTKEY)
                {
                    var handler = HotkeyPressed;
                    if (handler != null)
                        handler(m.WParam.ToInt32());
                }

                base.WndProc(ref m);
            }

            public void Dispose()
            {
                if (Handle != IntPtr.Zero)
                    DestroyHandle();
            }
        }

        private static bool IsForegroundWordWindow()
        {
            try
            {
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd == IntPtr.Zero) return false;
                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                return pid == _wordProcId;
            }
            catch { return true; }
        }
    }
}
