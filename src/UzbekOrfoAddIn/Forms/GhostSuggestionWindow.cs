using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>Per-pixel-alpha, click-through inline completion tail.</summary>
    internal sealed class GhostSuggestionWindow : Form
    {
        private const int WsExLayered = 0x00080000;
        private const int WsExTransparent = 0x00000020;
        private const int WsExNoActivate = 0x08000000;
        private const int WsExToolWindow = 0x00000080;
        private const int UlwAlpha = 0x00000002;
        private const byte AcSrcAlpha = 0x01;

        [StructLayout(LayoutKind.Sequential)]
        private struct NativePoint
        {
            internal int X;
            internal int Y;
            internal NativePoint(int x, int y) { X = x; Y = y; }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeSize
        {
            internal int Width;
            internal int Height;
            internal NativeSize(int width, int height) { Width = width; Height = height; }
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct BlendFunction
        {
            internal byte BlendOp;
            internal byte BlendFlags;
            internal byte SourceConstantAlpha;
            internal byte AlphaFormat;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr window);
        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr window, IntPtr dc);
        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr dc);
        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr dc);
        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr dc, IntPtr value);
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr value);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UpdateLayeredWindow(IntPtr window, IntPtr destinationDc,
            ref NativePoint destination, ref NativeSize size, IntPtr sourceDc,
            ref NativePoint source, int colorKey, ref BlendFunction blend, int flags);

        internal string SuggestionTail { get; private set; }
        internal Rectangle PresentationBounds { get; private set; }
        private string _fontName;
        private float _fontSizePoints;
        private FontStyle _fontStyle;
        private int _dpi;
        private bool _darkBackground;

        protected override bool ShowWithoutActivation => true;
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WsExLayered | WsExTransparent | WsExNoActivate | WsExToolWindow;
                return cp;
            }
        }

        internal GhostSuggestionWindow()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
            AccessibleRole = AccessibleRole.StaticText;
            AccessibleName = "MatnAI кулранг таклифи";
        }

        internal bool Present(string tail, OverlayAnchor anchor, IWin32Window owner)
        {
            if (string.IsNullOrEmpty(tail) || anchor == null)
            {
                Hide();
                return false;
            }

            bool sameVisual = Visible && string.Equals(tail, SuggestionTail, StringComparison.Ordinal) &&
                string.Equals(anchor.FontName, _fontName, StringComparison.OrdinalIgnoreCase) &&
                Math.Abs(anchor.FontSizePoints - _fontSizePoints) < 0.01f &&
                anchor.FontStyle == _fontStyle && anchor.Dpi == _dpi &&
                anchor.DarkBackground == _darkBackground;
            if (sameVisual)
            {
                Rectangle moved = OverlayPositioner.PlaceGhost(anchor, PresentationBounds.Size);
                if (moved.IsEmpty)
                {
                    Hide();
                    return false;
                }
                OverlayPositioner.MovePhysical(this, moved, anchor.OwnerHandle);
                PresentationBounds = moved;
                return true;
            }

            using (Font font = CreateFont(anchor))
            using (var measuring = new Bitmap(1, 1, PixelFormat.Format32bppArgb))
            using (Graphics graphics = Graphics.FromImage(measuring))
            using (var format = (StringFormat)StringFormat.GenericTypographic.Clone())
            {
                graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                format.FormatFlags |= StringFormatFlags.NoWrap | StringFormatFlags.MeasureTrailingSpaces;
                SizeF measured = graphics.MeasureString(tail, font, int.MaxValue, format);
                int width = Math.Max(2, (int)Math.Ceiling(measured.Width) + ScreenGeometry.Scale(5, anchor.Dpi));
                int height = Math.Max(anchor.CaretBounds.Height,
                    (int)Math.Ceiling(font.GetHeight(graphics)) + ScreenGeometry.Scale(4, anchor.Dpi));
                Rectangle bounds = OverlayPositioner.PlaceGhost(anchor, new Size(width, height));
                if (bounds.IsEmpty)
                {
                    Hide();
                    return false;
                }

                using (new DpiLayout.Context(DpiLayout.WindowContext(anchor.OwnerHandle)))
                {
                    if (!IsHandleCreated) { var unused = Handle; }
                    if (!Visible)
                    {
                        if (owner == null) Show(); else Show(owner);
                    }
                }

                Color color = anchor.DarkBackground
                    ? Color.FromArgb(220, 166, 180, 196)
                    : Color.FromArgb(220, 112, 128, 144);
                if (!RenderLayer(tail, font, format, color, bounds))
                {
                    Hide();
                    return false;
                }
                SuggestionTail = tail;
                PresentationBounds = bounds;
                _fontName = anchor.FontName;
                _fontSizePoints = anchor.FontSizePoints;
                _fontStyle = anchor.FontStyle;
                _dpi = anchor.Dpi;
                _darkBackground = anchor.DarkBackground;
                AccessibleName = "MatnAI таклифи: " + tail;
                AccessibleDescription = "Tab тугмаси билан қабул қилинг";
                AccessibilityNotifyClients(AccessibleEvents.NameChange, -1);
                return true;
            }
        }

        private bool RenderLayer(string tail, Font font, StringFormat format, Color color, Rectangle bounds)
        {
            // UpdateLayeredWindow requires premultiplied alpha for clean antialiased edges.
            using (var bitmap = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppPArgb))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                using (var brush = new SolidBrush(color))
                {
                    graphics.Clear(Color.Transparent);
                    graphics.CompositingMode = CompositingMode.SourceOver;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                    float textHeight = font.GetHeight(graphics);
                    float y = Math.Max(0f, (bounds.Height - textHeight) / 2f - 1f);
                    graphics.DrawString(tail, font, brush, new PointF(0, y), format);
                }

                IntPtr screenDc = GetDC(IntPtr.Zero);
                IntPtr memoryDc = IntPtr.Zero;
                IntPtr bitmapHandle = IntPtr.Zero;
                IntPtr previous = IntPtr.Zero;
                try
                {
                    if (screenDc == IntPtr.Zero) return false;
                    memoryDc = CreateCompatibleDC(screenDc);
                    if (memoryDc == IntPtr.Zero) return false;
                    bitmapHandle = bitmap.GetHbitmap(Color.FromArgb(0));
                    previous = SelectObject(memoryDc, bitmapHandle);
                    var destination = new NativePoint(bounds.X, bounds.Y);
                    var source = new NativePoint(0, 0);
                    var size = new NativeSize(bounds.Width, bounds.Height);
                    var blend = new BlendFunction
                    {
                        BlendOp = 0,
                        BlendFlags = 0,
                        SourceConstantAlpha = 255,
                        AlphaFormat = AcSrcAlpha
                    };
                    using (new DpiLayout.Context(DpiLayout.WindowContext(Handle)))
                    {
                        if (!UpdateLayeredWindow(Handle, screenDc, ref destination, ref size,
                            memoryDc, ref source, 0, ref blend, UlwAlpha))
                            throw new Win32Exception(Marshal.GetLastWin32Error());
                    }
                    return true;
                }
                catch (Win32Exception)
                {
                    return false;
                }
                finally
                {
                    if (previous != IntPtr.Zero && memoryDc != IntPtr.Zero) SelectObject(memoryDc, previous);
                    if (bitmapHandle != IntPtr.Zero) DeleteObject(bitmapHandle);
                    if (memoryDc != IntPtr.Zero) DeleteDC(memoryDc);
                    if (screenDc != IntPtr.Zero) ReleaseDC(IntPtr.Zero, screenDc);
                }
            }
        }

        private static Font CreateFont(OverlayAnchor anchor)
        {
            float pixelSize = Math.Max(8f, anchor.FontSizePoints * anchor.Dpi / 72f);
            try { return new Font(anchor.FontName, pixelSize, anchor.FontStyle, GraphicsUnit.Pixel); }
            catch (ArgumentException) { return new Font("Segoe UI", pixelSize, FontStyle.Regular, GraphicsUnit.Pixel); }
        }

        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0x21) // WM_MOUSEACTIVATE
            {
                message.Result = new IntPtr(3); // MA_NOACTIVATE
                return;
            }
            if (message.Msg == 0x84) // WM_NCHITTEST
            {
                message.Result = new IntPtr(-1); // HTTRANSPARENT
                return;
            }
            base.WndProc(ref message);
        }
    }
}
