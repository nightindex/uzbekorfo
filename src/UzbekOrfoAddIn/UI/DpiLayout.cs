using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.UI
{
    internal static class DpiLayout
    {
        internal static int Pixels(Control control, int value) =>
            ScreenGeometry.Scale(value, (control.FindForm() as DpiForm)?.LayoutDpi ?? 96);

        internal static Font Font(Control control, Font designFont) =>
            (control.FindForm() as DpiForm)?.GetDpiFont(designFont) ?? designFont;

        [DllImport("user32.dll")]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDpiAwarenessContext(IntPtr handle);

        internal static IntPtr WindowContext(Control control)
        {
            try { return GetWindowDpiAwarenessContext(control.Handle); }
            catch (EntryPointNotFoundException) { return IntPtr.Zero; }
        }

        internal static Screen ActiveScreen() => Screen.FromHandle(GetForegroundWindow());

        internal static int WindowDpi(Control control)
        {
            try
            {
                uint dpi = GetDpiForWindow(control.Handle);
                if (dpi > 0) return (int)dpi;
            }
            catch (EntryPointNotFoundException) { }
            using (var graphics = control.CreateGraphics())
                return (int)graphics.DpiX;
        }

        // Office owns the process DPI policy. Only scope creation of our HWND;
        // restore the calling thread before executing any other Office code.
        internal sealed class Context : IDisposable
        {
            private readonly IntPtr previous;
            internal Context() : this(new IntPtr(-4)) { }

            internal Context(IntPtr context)
            {
                if (context == IntPtr.Zero) return;
                try
                {
                    previous = SetThreadDpiAwarenessContext(context);
                    if (previous == IntPtr.Zero)
                        previous = SetThreadDpiAwarenessContext(new IntPtr(-3));
                }
                catch (EntryPointNotFoundException) { }
            }
            public void Dispose()
            {
                if (previous != IntPtr.Zero) SetThreadDpiAwarenessContext(previous);
            }
        }
    }
}
