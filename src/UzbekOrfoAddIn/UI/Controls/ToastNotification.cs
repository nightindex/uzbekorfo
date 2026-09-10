using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.UI.Controls
{
    /// <summary>
    /// A non-blocking toast notification that slides in from the bottom-right corner.
    /// Auto-dismisses after a configurable duration with a progress bar.
    /// 
    /// Usage: ToastNotification.Show("Текширув тугади", "14 та хато топилди", ToastType.Success);
    /// </summary>
    public class ToastNotification : Form
    {
        public enum ToastType { Success, Error, Warning, Info }

        private readonly Timer _dismissTimer;
        private readonly Timer _progressTimer;
        private float _progress = 1.0f;
        private readonly int _durationMs;
        private readonly string _title;
        private readonly string _message;
        private readonly ToastType _type;
        private bool _hovered;

        private static ToastNotification _current;

        private const int TOAST_WIDTH = 400;
        private const int TOAST_HEIGHT = 90;
        private const int MARGIN = 20;
        private const int ACCENT_WIDTH = 5;

        // Cached icon font to avoid GDI leak in OnPaint
        private static readonly Font _iconFont = new Font("Segoe UI", 20f);

        private ToastNotification(string title, string message, ToastType type, int durationMs)
        {
            _title = title;
            _message = message;
            _type = type;
            _durationMs = durationMs;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(TOAST_WIDTH, TOAST_HEIGHT);
            DoubleBuffered = true;
            BackColor = ThemeManager.SurfaceElevated;

            // Position: bottom-right of screen
            var screen = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(
                screen.Right - Width - MARGIN,
                screen.Bottom - Height - MARGIN);

            // Rounded corners
            using (var path = CreateRoundedRect(new Rectangle(0, 0, Width, Height), ThemeManager.RadiusMD))
            {
                Region = new Region(path);
            }

            // Auto-dismiss timer
            _dismissTimer = new Timer { Interval = durationMs };
            _dismissTimer.Tick += async (s, e) =>
            {
                _dismissTimer.Stop();
                _progressTimer?.Stop();
                await AnimationHelper.FadeOut(this, 200);
                Close();
            };

            // Progress bar animation
            _progressTimer = new Timer { Interval = 30 };
            _progressTimer.Tick += (s, e) =>
            {
                if (_hovered) return; // pause when hovering
                _progress -= 30f / _durationMs;
                if (_progress < 0) _progress = 0;
                Invalidate(new Rectangle(0, Height - 3, Width, 3));
            };
        }

        // =====================================================================
        //  PUBLIC API
        // =====================================================================

        /// <summary>
        /// Shows a toast notification. If a previous toast is visible, it is replaced.
        /// </summary>
        public static void Show(string title, string message,
            ToastType type = ToastType.Info, int durationMs = 3000)
        {
            try
            {
                // Close previous toast
                if (_current != null && !_current.IsDisposed)
                {
                    _current.Close();
                    _current.Dispose();
                }

                _current = new ToastNotification(title, message, type, durationMs);
                _current.Opacity = 0;
                _current.Show();

                // Fade in
                var fadeTimer = new Timer { Interval = 15 };
                int fadeStep = 0;
                fadeTimer.Tick += (s, e) =>
                {
                    fadeStep++;
                    if (_current == null || _current.IsDisposed)
                    {
                        fadeTimer.Stop();
                        fadeTimer.Dispose();
                        return;
                    }
                    _current.Opacity = Math.Min(1.0, fadeStep * 0.08);
                    if (_current.Opacity >= 1.0)
                    {
                        fadeTimer.Stop();
                        fadeTimer.Dispose();
                        _current._dismissTimer.Start();
                        _current._progressTimer.Start();
                    }
                };
                fadeTimer.Start();
            }
            catch
            {
                // Toast should never crash the main application
            }
        }

        /// <summary>Shorthand for success toast.</summary>
        public static void Success(string title, string message = "") =>
            Show(title, message, ToastType.Success);

        /// <summary>Shorthand for error toast.</summary>
        public static void ShowError(string title, string message = "") =>
            Show(title, message, ToastType.Error, 5000);

        /// <summary>Shorthand for warning toast.</summary>
        public static void ShowWarning(string title, string message = "") =>
            Show(title, message, ToastType.Warning, 4000);

        /// <summary>Shorthand for info toast.</summary>
        public static void ShowInfo(string title, string message = "") =>
            Show(title, message, ToastType.Info);

        // =====================================================================
        //  PAINTING
        // =====================================================================

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Background
            using (var brush = new SolidBrush(ThemeManager.SurfaceElevated))
            {
                g.FillRectangle(brush, ClientRectangle);
            }

            // Border
            using (var path = CreateRoundedRect(new Rectangle(0, 0, Width - 1, Height - 1), ThemeManager.RadiusMD))
            using (var pen = new Pen(ThemeManager.Border))
            {
                g.DrawPath(pen, path);
            }

            // Accent bar on left
            Color accentColor = GetAccentColor();
            using (var brush = new SolidBrush(accentColor))
            {
                g.FillRectangle(brush, 0, 8, ACCENT_WIDTH, Height - 16);
            }

            // Icon
            string icon = GetIcon();
            using (var brush = new SolidBrush(accentColor))
            {
                g.DrawString(icon, _iconFont, brush, 14, 16);
            }

            // Title
            using (var brush = new SolidBrush(ThemeManager.TextPrimary))
            {
                g.DrawString(_title, ThemeManager.FontLGBold, brush, 52, 16);
            }

            // Message
            if (!string.IsNullOrEmpty(_message))
            {
                using (var brush = new SolidBrush(ThemeManager.TextSecondary))
                {
                    var msgRect = new RectangleF(52, 42, Width - 68, Height - 50);
                    g.DrawString(_message, ThemeManager.FontBase, brush, msgRect);
                }
            }

            // Progress bar at bottom
            int barHeight = 3;
            int barY = Height - barHeight;
            using (var brush = new SolidBrush(ThemeManager.Border))
            {
                g.FillRectangle(brush, 0, barY, Width, barHeight);
            }
            int progressWidth = (int)(Width * _progress);
            using (var brush = new SolidBrush(accentColor))
            {
                g.FillRectangle(brush, 0, barY, progressWidth, barHeight);
            }
        }

        private Color GetAccentColor()
        {
            switch (_type)
            {
                case ToastType.Success: return ThemeManager.Success;
                case ToastType.Error: return ThemeManager.Error;
                case ToastType.Warning: return ThemeManager.Warning;
                case ToastType.Info: return ThemeManager.Primary;
                default: return ThemeManager.Primary;
            }
        }

        private string GetIcon()
        {
            switch (_type)
            {
                case ToastType.Success: return "✓";
                case ToastType.Error: return "✕";
                case ToastType.Warning: return "⚠";
                case ToastType.Info: return "ℹ";
                default: return "ℹ";
            }
        }

        // =====================================================================
        //  MOUSE — click to dismiss, hover to pause
        // =====================================================================

        protected override void OnMouseEnter(EventArgs e)
        {
            _hovered = true;
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hovered = false;
            base.OnMouseLeave(e);
        }

        protected override void OnClick(EventArgs e)
        {
            _dismissTimer.Stop();
            _progressTimer.Stop();
            Close();
        }

        // =====================================================================
        //  SHADOW
        // =====================================================================

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ClassStyle |= 0x00020000; // CS_DROPSHADOW
                return cp;
            }
        }

        // =====================================================================
        //  HELPERS
        // =====================================================================

        private static GraphicsPath CreateRoundedRect(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _dismissTimer?.Dispose();
                _progressTimer?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
