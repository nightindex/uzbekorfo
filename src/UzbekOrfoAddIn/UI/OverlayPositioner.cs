using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.UI
{
    /// <summary>Physical-pixel anchor shared by MatnAI overlays.</summary>
    internal sealed class OverlayAnchor
    {
        internal Rectangle CaretBounds { get; }
        internal Rectangle WorkArea { get; }
        internal IntPtr OwnerHandle { get; }
        internal int Dpi { get; }
        internal string FontName { get; }
        internal float FontSizePoints { get; }
        internal FontStyle FontStyle { get; }
        internal bool DarkBackground { get; }

        internal OverlayAnchor(Rectangle caretBounds, Rectangle workArea, IntPtr ownerHandle,
            int dpi, string fontName, float fontSizePoints, FontStyle fontStyle, bool darkBackground)
        {
            CaretBounds = new Rectangle(caretBounds.X, caretBounds.Y,
                Math.Max(1, caretBounds.Width), Math.Max(1, caretBounds.Height));
            WorkArea = workArea;
            OwnerHandle = ownerHandle;
            Dpi = Math.Max(96, dpi);
            FontName = string.IsNullOrWhiteSpace(fontName) ? "Segoe UI" : fontName;
            FontSizePoints = fontSizePoints >= 6f && fontSizePoints <= 96f ? fontSizePoints : 11f;
            FontStyle = fontStyle;
            DarkBackground = darkBackground;
        }
    }

    /// <summary>
    /// Places non-activating overlays using physical screen pixels. This avoids
    /// WinForms coordinate virtualization when Word spans differently scaled monitors.
    /// </summary>
    internal static class OverlayPositioner
    {
        private const uint MonitorDefaultToNearest = 2;
        private const uint SwpNoZOrder = 0x0004;
        private const uint SwpNoActivate = 0x0010;
        private const uint SwpNoOwnerZOrder = 0x0200;

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect { internal int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct MonitorInfo
        {
            internal int Size;
            internal NativeRect Monitor;
            internal NativeRect Work;
            internal uint Flags;
        }

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr window);
        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromRect(ref NativeRect rect, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr window, IntPtr insertAfter,
            int x, int y, int width, int height, uint flags);
        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr window);
        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr window, IntPtr dc);
        [DllImport("gdi32.dll")]
        private static extern uint GetPixel(IntPtr dc, int x, int y);

        internal static OverlayAnchor CreateAnchor(Rectangle caret, IntPtr ownerHandle,
            string fontName, float fontSizePoints, FontStyle fontStyle)
        {
            return new OverlayAnchor(caret, GetWorkArea(caret), ownerHandle,
                GetOwnerDpi(ownerHandle), fontName, fontSizePoints, fontStyle,
                IsDarkAtCaret(caret));
        }

        internal static Rectangle PlacePopup(OverlayAnchor anchor, Size size)
        {
            if (anchor == null || size.Width <= 0 || size.Height <= 0) return Rectangle.Empty;
            Rectangle area = anchor.WorkArea;
            int gap = ScreenGeometry.Scale(4, anchor.Dpi);
            int width = Math.Min(size.Width, area.Width);
            int height = Math.Min(size.Height, area.Height);
            int x = Math.Max(area.Left, Math.Min(anchor.CaretBounds.Left, area.Right - width));
            int below = anchor.CaretBounds.Bottom + gap;
            int above = anchor.CaretBounds.Top - gap - height;
            int y = below + height <= area.Bottom ? below : Math.Max(area.Top, above);
            return new Rectangle(x, y, width, height);
        }

        internal static Rectangle PlaceGhost(OverlayAnchor anchor, Size size)
        {
            if (anchor == null || size.Width <= 0 || size.Height <= 0) return Rectangle.Empty;
            Rectangle area = anchor.WorkArea;
            int x = anchor.CaretBounds.Right + ScreenGeometry.Scale(1, anchor.Dpi);
            int y = anchor.CaretBounds.Top + (anchor.CaretBounds.Height - size.Height) / 2;
            if (x < area.Left || x + size.Width > area.Right || size.Height > area.Height)
                return Rectangle.Empty; // A ghost must remain inline; never move it to a misleading position.
            y = Math.Max(area.Top, Math.Min(y, area.Bottom - size.Height));
            return new Rectangle(x, y, size.Width, size.Height);
        }

        internal static void MovePhysical(Form overlay, Rectangle bounds, IntPtr ownerHandle)
        {
            if (overlay == null || overlay.IsDisposed || bounds.IsEmpty) return;
            IntPtr contextHandle = overlay.IsHandleCreated ? overlay.Handle : ownerHandle;
            using (new DpiLayout.Context(DpiLayout.WindowContext(contextHandle)))
            {
                SetWindowPos(overlay.Handle, IntPtr.Zero, bounds.X, bounds.Y,
                    bounds.Width, bounds.Height, SwpNoZOrder | SwpNoActivate | SwpNoOwnerZOrder);
            }
        }

        private static int GetOwnerDpi(IntPtr ownerHandle)
        {
            if (ownerHandle == IntPtr.Zero) return 96;
            try
            {
                uint dpi = GetDpiForWindow(ownerHandle);
                return dpi >= 96 && dpi <= 768 ? (int)dpi : 96;
            }
            catch (EntryPointNotFoundException) { return 96; }
        }

        private static Rectangle GetWorkArea(Rectangle bounds)
        {
            var rect = new NativeRect
            {
                Left = bounds.Left,
                Top = bounds.Top,
                Right = Math.Max(bounds.Left + 1, bounds.Right),
                Bottom = Math.Max(bounds.Top + 1, bounds.Bottom)
            };
            IntPtr monitor = MonitorFromRect(ref rect, MonitorDefaultToNearest);
            var info = new MonitorInfo { Size = Marshal.SizeOf(typeof(MonitorInfo)) };
            if (monitor != IntPtr.Zero && GetMonitorInfo(monitor, ref info))
                return Rectangle.FromLTRB(info.Work.Left, info.Work.Top, info.Work.Right, info.Work.Bottom);
            return Screen.FromRectangle(bounds).WorkingArea;
        }

        private static bool IsDarkAtCaret(Rectangle caret)
        {
            IntPtr dc = GetDC(IntPtr.Zero);
            if (dc == IntPtr.Zero) return ThemeManager.IsDarkTheme;
            try
            {
                int sampleX = caret.Right + 3;
                int sampleY = caret.Top + Math.Max(1, caret.Height / 2);
                uint value = GetPixel(dc, sampleX, sampleY);
                if (value == 0xffffffff) return ThemeManager.IsDarkTheme;
                int red = (int)(value & 0xff);
                int green = (int)((value >> 8) & 0xff);
                int blue = (int)((value >> 16) & 0xff);
                return red * 299 + green * 587 + blue * 114 < 128000;
            }
            finally { ReleaseDC(IntPtr.Zero, dc); }
        }
    }
}
