using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.UI.Controls
{
    /// <summary>
    /// A modern iOS-style toggle switch control with smooth animation.
    /// Shows On/Off state with sliding thumb and color transition.
    /// </summary>
    public class ModernToggle : Control
    {
        private bool _isOn;
        private float _thumbPosition; // 0.0 (off) to 1.0 (on) for animation
        private Timer _animTimer;

        /// <summary>
        /// Whether the toggle is in the ON state.
        /// </summary>
        public bool IsOn
        {
            get => _isOn;
            set
            {
                if (_isOn == value) return;
                _isOn = value;
                StartAnimation();
                Toggled?.Invoke(this, EventArgs.Empty);
            }
        }

        /// <summary>
        /// Text shown next to the toggle when ON.
        /// </summary>
        public string OnText { get; set; } = "Ёқ";

        /// <summary>
        /// Text shown next to the toggle when OFF.
        /// </summary>
        public string OffText { get; set; } = "Ўч";

        /// <summary>
        /// Whether to show text label next to the toggle.
        /// </summary>
        public bool ShowLabel { get; set; } = true;

        /// <summary>
        /// Fires when the toggle state changes.
        /// </summary>
        public event EventHandler Toggled;

        public ModernToggle()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw, true);

            Size = new Size(ShowLabel ? 90 : 48, 26);
            Cursor = Cursors.Hand;

            _animTimer = new Timer { Interval = 16 }; // ~60fps
            _animTimer.Tick += OnAnimTick;
        }

        // =====================================================================
        //  ANIMATION
        // =====================================================================

        private void StartAnimation()
        {
            _animTimer.Start();
        }

        private void OnAnimTick(object sender, EventArgs e)
        {
            float target = _isOn ? 1.0f : 0.0f;
            float speed = 0.15f;

            _thumbPosition += (_thumbPosition < target ? speed : -speed);

            if (Math.Abs(_thumbPosition - target) < speed)
            {
                _thumbPosition = target;
                _animTimer.Stop();
            }

            Invalidate();
        }

        // =====================================================================
        //  PAINTING
        // =====================================================================

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            int trackWidth = 44;
            int trackHeight = 24;
            int thumbSize = 18;
            int thumbPad = 3;

            // Track
            var trackRect = new Rectangle(0, (Height - trackHeight) / 2, trackWidth, trackHeight);
            Color trackColor = InterpolateColor(ThemeManager.Border, ThemeManager.Primary, _thumbPosition);
            using (var path = CreatePillPath(trackRect))
            using (var brush = new SolidBrush(trackColor))
            {
                g.FillPath(brush, path);
            }

            // Thumb
            float thumbMinX = trackRect.X + thumbPad;
            float thumbMaxX = trackRect.Right - thumbSize - thumbPad;
            float thumbX = thumbMinX + (thumbMaxX - thumbMinX) * _thumbPosition;
            int thumbY = trackRect.Y + (trackHeight - thumbSize) / 2;

            var thumbRect = new RectangleF(thumbX, thumbY, thumbSize, thumbSize);
            using (var brush = new SolidBrush(Color.White))
            {
                g.FillEllipse(brush, thumbRect);
            }

            // Thumb shadow
            using (var pen = new Pen(Color.FromArgb(30, 0, 0, 0), 1))
            {
                g.DrawEllipse(pen, thumbRect);
            }

            // Label text
            if (ShowLabel)
            {
                string label = _isOn ? OnText : OffText;
                Color labelColor = _isOn ? ThemeManager.Primary : ThemeManager.TextSecondary;
                using (var brush = new SolidBrush(labelColor))
                {
                    var textRect = new RectangleF(trackWidth + 8, 0, Width - trackWidth - 8, Height);
                    var sf = new StringFormat
                    {
                        Alignment = StringAlignment.Near,
                        LineAlignment = StringAlignment.Center
                    };
                    g.DrawString(label, ThemeManager.FontBase, brush, textRect, sf);
                }
            }
        }

        // =====================================================================
        //  MOUSE
        // =====================================================================

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            IsOn = !IsOn;
        }

        // =====================================================================
        //  HELPERS
        // =====================================================================

        private static GraphicsPath CreatePillPath(Rectangle rect)
        {
            var path = new GraphicsPath();
            int radius = rect.Height;
            path.AddArc(rect.X, rect.Y, radius, radius, 90, 180);
            path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 180);
            path.CloseFigure();
            return path;
        }

        private static Color InterpolateColor(Color from, Color to, float t)
        {
            return Color.FromArgb(
                (int)(from.R + (to.R - from.R) * t),
                (int)(from.G + (to.G - from.G) * t),
                (int)(from.B + (to.B - from.B) * t));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _animTimer?.Stop();
                _animTimer?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
