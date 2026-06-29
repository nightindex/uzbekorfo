using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.UI.Controls
{
    /// <summary>
    /// A modern card panel with elevated surface, rounded corners, and optional header.
    /// Use as a container for grouping related controls within a ModernForm.
    /// </summary>
    public class ModernCard : Panel
    {
        private string _header;
        private int _cornerRadius = 8;

        /// <summary>
        /// Optional header text shown at the top of the card.
        /// </summary>
        public string Header
        {
            get => _header;
            set { _header = value; UpdatePadding(); Invalidate(); }
        }

        /// <summary>
        /// Corner radius for the card border.
        /// </summary>
        public int CornerRadius
        {
            get => _cornerRadius;
            set { _cornerRadius = value; Invalidate(); }
        }

        public ModernCard()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw, true);

            BackColor = ThemeManager.Surface;
            UpdatePadding();
        }

        private void UpdatePadding()
        {
            int topPad = string.IsNullOrEmpty(_header)
                ? ThemeManager.SpaceLG
                : ThemeManager.SpaceLG + 28; // header height + spacing
            Padding = new Padding(ThemeManager.SpaceLG, topPad, ThemeManager.SpaceLG, ThemeManager.SpaceLG);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);

            // Background
            using (var path = CreateRoundedRect(rect, _cornerRadius))
            using (var brush = new SolidBrush(ThemeManager.Surface))
            {
                g.FillPath(brush, path);
            }

            // Border
            using (var path = CreateRoundedRect(rect, _cornerRadius))
            using (var pen = new Pen(ThemeManager.Border))
            {
                g.DrawPath(pen, path);
            }

            // Header
            if (!string.IsNullOrEmpty(_header))
            {
                using (var brush = new SolidBrush(ThemeManager.TextSecondary))
                {
                    g.DrawString(_header, ThemeManager.FontBaseBold, brush,
                        ThemeManager.SpaceLG, ThemeManager.SpaceMD);
                }

                // Header separator line
                int lineY = ThemeManager.SpaceMD + 22;
                using (var pen = new Pen(ThemeManager.Border))
                {
                    g.DrawLine(pen, ThemeManager.SpaceMD, lineY, Width - ThemeManager.SpaceMD, lineY);
                }
            }
        }

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
    }
}
