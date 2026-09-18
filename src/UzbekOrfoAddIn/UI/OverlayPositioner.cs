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
        internal int TextRight { get; set; } = int.MaxValue;
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
            FontSizePoints = fontSizePoints >= 1f && fontSizePoints <= 600f ? fontSizePoints : 11f;
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

        [StructLayout(LayoutKind.Sequential)]
        private struct GuiThreadInfo
        {
            internal int Size;
            internal uint Flags;
            internal IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
            internal NativeRect CaretRect;
        }

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

        [DllImport("user32.dll")]
        private static extern bool GetGUIThreadInfo(uint thread, ref GuiThreadInfo info);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr window, out NativeRect rect);
        [DllImport("user32.dll")]
        private static extern int MapWindowPoints(IntPtr from, IntPtr to, ref NativeRect rect, uint count);

        internal static bool TryGetNativeCaret(IntPtr owner, out Rectangle caret)
        {
            caret = Rectangle.Empty;
            if (owner == IntPtr.Zero || GetForegroundWindow() != owner) return false;
            var info = new GuiThreadInfo { Size = Marshal.SizeOf(typeof(GuiThreadInfo)) };
            if (!GetGUIThreadInfo(0, ref info) || info.Active != owner || info.Caret == IntPtr.Zero ||
                info.Focus != info.Caret || !IsChild(owner, info.Caret) ||
                info.CaretRect.Bottom <= info.CaretRect.Top) return false;
            // rcCaret uses the caret HWND's logical coordinates. Map in that HWND's
            // DPI context so it agrees with Word's screen coordinates on each monitor.
            using (new DpiLayout.Context(DpiLayout.WindowContext(info.Caret)))
                MapWindowPoints(info.Caret, IntPtr.Zero, ref info.CaretRect, 2);
            caret = Rectangle.FromLTRB(info.CaretRect.Left, info.CaretRect.Top,
                Math.Max(info.CaretRect.Left + 1, info.CaretRect.Right), info.CaretRect.Bottom);
            return true;
        }

        internal static OverlayAnchor CreateAnchor(Rectangle caret, IntPtr ownerHandle,
            string fontName, float fontSizePoints, FontStyle fontStyle)
        {
            Rectangle area = GetWorkArea(caret);
            var info = new GuiThreadInfo { Size = Marshal.SizeOf(typeof(GuiThreadInfo)) };
            NativeRect client;
            if (GetGUIThreadInfo(0, ref info) && info.Active == ownerHandle &&
                IsChild(ownerHandle, info.Focus) && GetClientRect(info.Focus, out client))
            {
                MapWindowPoints(info.Focus, IntPtr.Zero, ref client, 2);
                area = Rectangle.Intersect(area, Rectangle.FromLTRB(client.Left, client.Top, client.Right, client.Bottom));
            }
            return new OverlayAnchor(caret, area, ownerHandle,
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

        internal static OverlayAnchor WithTextHeight(OverlayAnchor anchor)
        {
            // GetPoint includes paragraph spacing in a character's rectangle. Use
            // the font's cell height at the character top when no native caret exists.
            using (var font = new Font(anchor.FontName, anchor.FontSizePoints * anchor.Dpi / 72f,
                anchor.FontStyle, GraphicsUnit.Pixel))
            {
                float height = font.Size * (font.FontFamily.GetCellAscent(font.Style) +
                    font.FontFamily.GetCellDescent(font.Style)) / font.FontFamily.GetEmHeight(font.Style);
                var caret = new Rectangle(anchor.CaretBounds.Left,
                    anchor.CaretBounds.Top, 1, (int)Math.Round(height));
                return new OverlayAnchor(caret, anchor.WorkArea, anchor.OwnerHandle, anchor.Dpi,
                    anchor.FontName, anchor.FontSizePoints, anchor.FontStyle, anchor.DarkBackground) { TextRight = anchor.TextRight };
            }
        }

        internal static Rectangle PlaceGhost(OverlayAnchor anchor, Size size, float textCellHeight = 0f)
        {
            if (anchor == null || size.Width <= 0 || size.Height <= 0) return Rectangle.Empty;
            Rectangle area = anchor.WorkArea;
            int x = anchor.CaretBounds.Right;
            int y = textCellHeight > 0f
                ? (int)Math.Round(anchor.CaretBounds.Bottom - textCellHeight)
                : anchor.CaretBounds.Top;
            if (x < area.Left || x + size.Width > Math.Min(area.Right, anchor.TextRight) || y < area.Top || y + size.Height > area.Bottom)
                return Rectangle.Empty; // A ghost must remain inline; never move it to a misleading position.
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

        internal static void RaiseWithoutActivation(Form overlay, IntPtr ownerHandle)
        {
            // An owned non-activating layered HWND can remain behind Office after
            // Show/UpdateLayeredWindow. Correct its z-order without focusing it or
            // making it globally topmost. Never raise over a different application.
            if (overlay == null || overlay.IsDisposed || !overlay.IsHandleCreated ||
                ownerHandle == IntPtr.Zero || GetForegroundWindow() != ownerHandle) return;
            SetWindowPos(overlay.Handle, IntPtr.Zero /* HWND_TOP */, 0, 0, 0, 0,
                0x0001 /* SWP_NOSIZE */ | 0x0002 /* SWP_NOMOVE */ | SwpNoActivate | SwpNoOwnerZOrder);
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
                // Sampling to the right, halfway down the caret, samples OUR gray
                // glyphs. That feeds back into theme detection and alternates the
                // ghost color every polling tick. Sample above/left of the overlay.
                int sampleX = caret.Left - 4;
                int sampleY = caret.Top - 3;
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
