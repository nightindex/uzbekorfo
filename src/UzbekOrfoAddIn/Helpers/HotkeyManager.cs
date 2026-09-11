using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Handles shortcuts on Word's UI thread without reserving system-wide keys.
    /// Register/Unregister must run on that thread. Actions run after the hook returns.
    /// </summary>
    public static class HotkeyManager
    {
        [Flags]
        public enum Modifiers : byte { None = 0, Ctrl = 1, Shift = 2, Alt = 4 }

        public sealed class HotkeyDef
        {
            public Modifiers Mod { get; }
            public Keys Key { get; }
            public Action Action { get; }
            public string Label { get; }
            public HotkeyDef(Modifiers mod, Keys key, Action action, string label = "")
            {
                Mod = mod; Key = key; Action = action; Label = label;
            }
        }

        private delegate IntPtr KeyboardProc(int code, IntPtr key, IntPtr flags);
        private static readonly KeyboardProc HookCallback = OnKeyboard;
        private static readonly Dictionary<Keys, HotkeyDef> Hotkeys = new Dictionary<Keys, HotkeyDef>();
        private static readonly HashSet<Keys> ConsumedKeys = new HashSet<Keys>();
        private static IntPtr _hook;
        private static uint _threadId;
        private static Control _dispatcher;
        private static bool _executing;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr SetWindowsHookEx(int hook, KeyboardProc callback, IntPtr module, uint threadId);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr key, IntPtr flags);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int count);
        [DllImport("user32.dll")]
        private static extern bool IsWindowEnabled(IntPtr window);

        public static void Register(IEnumerable<HotkeyDef> hotkeys)
        {
            if (hotkeys == null) throw new ArgumentNullException(nameof(hotkeys));
            Unregister();
            try
            {
                foreach (var hotkey in hotkeys)
                {
                    if (hotkey == null || hotkey.Action == null)
                        throw new ArgumentException("Every shortcut requires an action.", nameof(hotkeys));
                    Keys chord = hotkey.Key & Keys.KeyCode;
                    if ((hotkey.Mod & Modifiers.Ctrl) != 0) chord |= Keys.Control;
                    if ((hotkey.Mod & Modifiers.Shift) != 0) chord |= Keys.Shift;
                    if ((hotkey.Mod & Modifiers.Alt) != 0) chord |= Keys.Alt;
                    if (Hotkeys.ContainsKey(chord))
                        throw new ArgumentException("Duplicate shortcut: " + hotkey.Label, nameof(hotkeys));
                    Hotkeys.Add(chord, hotkey);
                }
                _threadId = GetCurrentThreadId();
                _dispatcher = new Control();
                var handle = _dispatcher.Handle;
                _hook = SetWindowsHookEx(2 /* WH_KEYBOARD */, HookCallback, IntPtr.Zero, _threadId);
                if (_hook == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
                Logger.Info($"HotkeyManager: enabled {Hotkeys.Count} Word shortcuts.");
            }
            catch (Exception ex)
            {
                Unregister();
                Logger.Error("Word shortcut initialization failed", ex);
                MessageBox.Show("?????? ?????????? ???? ???????. ????????? ??????????? ???????????.",
                    "????? ????", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public static void Unregister()
        {
            if (_threadId != 0 && _threadId != GetCurrentThreadId())
                throw new InvalidOperationException("Shortcuts must be removed on their UI thread.");
            if (_hook != IntPtr.Zero)
            {
                if (!UnhookWindowsHookEx(_hook))
                    Logger.Warn("Unable to remove keyboard hook: " + Marshal.GetLastWin32Error());
                _hook = IntPtr.Zero;
            }
            Hotkeys.Clear();
            ConsumedKeys.Clear();
            _dispatcher?.Dispose();
            _dispatcher = null;
            _threadId = 0;
        }

        private static IntPtr OnKeyboard(int code, IntPtr keyValue, IntPtr flagsValue)
        {
            // HC_NOREMOVE and negative codes must pass through unchanged.
            if (code != 0) return CallNextHookEx(_hook, code, keyValue, flagsValue);
            try
            {
                var key = (Keys)keyValue.ToInt32();
                long flags = flagsValue.ToInt64();
                bool released = (flags & 0x80000000L) != 0;
                if (released)
                {
                    if (ConsumedKeys.Remove(key)) return new IntPtr(1);
                }
                else
                {
                    bool repeat = (flags & 0x40000000L) != 0;
                    if (repeat && ConsumedKeys.Contains(key)) return new IntPtr(1);
                    // Focus may have moved to another process before the previous
                    // key-up reached this thread. A new press starts a fresh cycle.
                    if (!repeat) ConsumedKeys.Remove(key);
                    if (!repeat && !_executing && IsForegroundWordWindow() &&
                        Hotkeys.TryGetValue(key | Control.ModifierKeys, out var hotkey))
                    {
                        IntPtr target = GetForegroundWindow();
                        var dispatcher = _dispatcher;
                        dispatcher.BeginInvoke((Action)(() =>
                        {
                            if (dispatcher != _dispatcher || _executing ||
                                GetForegroundWindow() != target || !IsForegroundWordWindow()) return;
                            _executing = true;
                            try { hotkey.Action(); }
                            catch (Exception ex) { Logger.Error("Shortcut failed: " + hotkey.Label, ex); }
                            finally { _executing = false; }
                        }));
                        ConsumedKeys.Add(key);
                        return new IntPtr(1);
                    }
                }
            }
            catch (Exception ex) { Logger.Error("Keyboard callback failed", ex); }
            return CallNextHookEx(_hook, code, keyValue, flagsValue);
        }

        private static bool IsForegroundWordWindow()
        {
            IntPtr window = GetForegroundWindow();
            if (window == IntPtr.Zero || !IsWindowEnabled(window) ||
                GetWindowThreadProcessId(window, out _) != _threadId) return false;
            var name = new StringBuilder(64);
            GetClassName(window, name, name.Capacity);
            // Excludes Word dialogs and the add-in's modeless forms.
            return string.Equals(name.ToString(), "OpusApp", StringComparison.Ordinal);
        }
    }
}
