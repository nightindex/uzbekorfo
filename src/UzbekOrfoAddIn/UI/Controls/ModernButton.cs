using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.UI.Controls
{
    /// <summary>
    /// A modern flat button with rounded corners, hover effects, and theme support.
    /// Supports Primary, Secondary, Danger, and Ghost styles.
    /// </summary>
    public class ModernButton : Control
    {
        private int Px(int value) => DpiLayout.Pixels(this, value);
        private Font UiFont(Font value) => DpiLayout.Font(this, value);

        // --- Enums ---
        public enum ButtonStyle { Primary, Secondary, Danger, Ghost }

        // --- Properties ---
        private ButtonStyle _style = ButtonStyle.Primary;
        public ButtonStyle Style
        {
            get => _style;
            set { _style = value; Invalidate(); }
        }

        private int _cornerRadius = 8;
        public int CornerRadius
        {
            get => _cornerRadius;
            set { _cornerRadius = value; Invalidate(); }
        }

        public Image Icon { get; set; }
        public ContentAlignment IconAlign { get; set; } = ContentAlignment.MiddleLeft;

        // --- State ---
        private bool _hovered;
        private bool _pressed;

        public ModernButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.SupportsTransparentBackColor |
                     ControlStyles.ResizeRedraw, true);

            BackColor = Color.Transparent;
            Size = new Size(140, 36);
            Font = ThemeManager.FontBase;
            Cursor = Cursors.Hand;
            TabStop = true;
            AccessibleRole = AccessibleRole.PushButton;
        }

        // =====================================================================
        //  COLOR RESOLUTION
        // =====================================================================

        private Color GetBackColor()
        {
            switch (_style)
            {
                case ButtonStyle.Primary:
                    return _pressed ? ThemeManager.PrimaryPressed :
                           _hovered ? ThemeManager.PrimaryHover : ThemeManager.Primary;
                case ButtonStyle.Secondary:
                    return _pressed ? ThemeManager.Border :
                           _hovered ? ThemeManager.ButtonSecondaryHover : ThemeManager.ButtonSecondary;
                case ButtonStyle.Danger:
                    return _pressed ? ThemeManager.ButtonDangerPressed :
                           _hovered ? ThemeManager.ButtonDangerHover : ThemeManager.ButtonDanger;
                case ButtonStyle.Ghost:
                    return _pressed ? ThemeManager.Border :
                           _hovered ? ThemeManager.SurfaceHover : Color.Transparent;
                default:
                    return ThemeManager.Primary;
            }
        }

        private Color GetForeColor()
        {
            switch (_style)
            {
                case ButtonStyle.Primary:
                case ButtonStyle.Danger:
                    return Color.White;
                case ButtonStyle.Secondary:
                case ButtonStyle.Ghost:
                    return ThemeManager.TextPrimary;
                default:
                    return Color.White;
            }
        }

        private Color GetBorderColor()
        {
            switch (_style)
            {
                case ButtonStyle.Secondary:
                    return _hovered ? ThemeManager.Primary : Color.Transparent;
                case ButtonStyle.Ghost:
                    return Color.Transparent;
                default:
                    return Color.Transparent;
            }
        }

        // =====================================================================
        //  PAINTING
        // =====================================================================

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Clear with parent background so rounded corners blend seamlessly
            var clearColor = GetEffectiveParentBackColor();
            g.Clear(clearColor);

            var rect = new Rectangle(1, 1, Width - 3, Height - 3);

            // Background
            using (var path = CreateRoundedRect(rect, Px(_cornerRadius)))
            using (var brush = new SolidBrush(GetBackColor()))
            {
                g.FillPath(brush, path);
            }

            // Border
            var borderColor = GetBorderColor();
            if (borderColor != Color.Transparent)
            {
                using (var path = CreateRoundedRect(rect, Px(_cornerRadius)))
                using (var pen = new Pen(borderColor, 1))
                {
                    g.DrawPath(pen, path);
                }
            }

            // Content layout
            var foreColor = Enabled ? GetForeColor() : ThemeManager.TextDisabled;
            var sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center,
                Trimming = StringTrimming.EllipsisCharacter
            };

            if (Icon != null)
            {
                int iconSize = Px(16);
                int totalWidth = iconSize + Px(6) + (int)g.MeasureString(Text, Font).Width;
                int startX = (Width - totalWidth) / 2;

                var iconRect = new Rectangle(startX, (Height - iconSize) / 2, iconSize, iconSize);
                g.DrawImage(Icon, iconRect);

                var textRect = new Rectangle(startX + iconSize + Px(6), 0, Width - startX - iconSize - Px(6), Height);
                sf.Alignment = StringAlignment.Near;
                using (var brush = new SolidBrush(foreColor))
                {
                    g.DrawString(Text, Font, brush, textRect, sf);
                }
            }
            else
            {
                using (var brush = new SolidBrush(foreColor))
                {
                    g.DrawString(Text, Font, brush, rect, sf);
                }
            }
            sf.Dispose();
            if (Focused && ShowFocusCues)
                ControlPaint.DrawFocusRectangle(g, Rectangle.Inflate(rect, -3, -3), foreColor, GetBackColor());
        }

        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

        // =====================================================================
        //  MOUSE EVENTS
        // =====================================================================

        protected override void OnMouseEnter(EventArgs e)
        {
            _hovered = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hovered = false;
            _pressed = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Focus();
                _pressed = true;
                Invalidate();
            }
            base.OnMouseDown(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            _pressed = false;
            Invalidate();
            base.OnMouseUp(e);
        }

        // =====================================================================
        //  KEYBOARD
        // =====================================================================

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
            {
                OnClick(EventArgs.Empty);
                e.Handled = true;
            }
            base.OnKeyDown(e);
        }

        protected override bool IsInputKey(Keys keyData)
        {
            if (keyData == Keys.Enter || keyData == Keys.Space)
                return true;
            return base.IsInputKey(keyData);
        }

        // =====================================================================
        //  HELPERS
        // =====================================================================

        private static GraphicsPath CreateRoundedRect(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = Math.Max(1, Math.Min(radius * 2, Math.Min(rect.Width, rect.Height)));
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        /// <summary>
        /// Walk up the parent chain to find the first non-transparent BackColor.
        /// </summary>
        private Color GetEffectiveParentBackColor()
        {
            var ctrl = Parent;
            while (ctrl != null)
            {
                if (ctrl.BackColor != Color.Transparent && ctrl.BackColor.A == 255)
                    return ctrl.BackColor;
                ctrl = ctrl.Parent;
            }
            return ThemeManager.Background;
        }
    }
}
