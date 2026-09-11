using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.UI
{
    /// <summary>
    /// Base class for all modern dialogs in Uzbek Orfo.
    /// Provides: custom title bar, rounded corners, shadow, fade animation,
    /// theme-aware colors, draggable window, and resize support.
    /// 
    /// Inherit from this class instead of System.Windows.Forms.Form.
    /// </summary>
    public class ModernForm : DpiForm
    {
        // --- Title bar ---
        private int TITLE_BAR_HEIGHT => Px(48);
        private int RESIZE_BORDER => Px(6);
        private bool _isDragging;
        private Point _dragStart;

        // --- Close/minimize buttons ---
        private Rectangle _closeBtnRect;
        private Rectangle _minBtnRect;
        private bool _closeHovered;
        private bool _minHovered;

        // --- Title spacer (stored for mouse forwarding) ---
        private Panel _titleSpacer;
        private Panel _contentViewport;
        private Panel _actionViewport;
        private Size _contentDesignSize;
        private Size _actionDesignSize;
        private int _actionMinimumWidth;
        private bool _layingOutViewports;

        protected virtual bool ReflowContent => false;

        protected override void OnLoad(EventArgs e)
        {
            using (new DpiLayout.Context())
            {
                if (_contentViewport == null)
                {
                    // Preserve the designed layout on small work areas. Both the
                    // body and the action buttons stay reachable through scrolling.
                    // A resizable body can shrink to the form's supported minimum.
                    // Keeping the preferred size here makes even ordinary resizing scroll.
                    _contentDesignSize = new Size(
                        Math.Max(1, MinimumSize.Width - (Width - ContentPanel.Width)),
                        Math.Max(1, MinimumSize.Height - (Height - ContentPanel.Height)));
                    _actionDesignSize = ActionBar.Size;
                    _actionMinimumWidth = GetMinimumActionBarWidth();
                    _contentViewport = WrapViewport(ContentPanel, DockStyle.Fill);
                    _actionViewport = WrapViewport(ActionBar, DockStyle.Bottom);
                    _actionViewport.Height = ActionBar.Height;
                    _contentViewport.Resize += (s, args) => LayoutViewports();
                    _actionViewport.Resize += (s, args) => LayoutViewports();
                }
                base.OnLoad(e);
                LayoutViewports();
            }
        }

        protected virtual int GetMinimumActionBarWidth()
        {
            var buttons = ActionBar.Controls.Cast<Control>().OfType<ModernButton>().ToArray();
            return buttons.Length == ActionBar.Controls.Count
                ? buttons.Sum(button => button.Width) + Math.Max(0, buttons.Length - 1) * 12 + 48
                : ActionBar.Width;
        }

        private Panel WrapViewport(Panel canvas, DockStyle dock)
        {
            int index = Controls.GetChildIndex(canvas);
            var viewport = new Panel { Dock = dock, AutoScroll = true, BackColor = canvas.BackColor,
                Margin = Padding.Empty, Padding = Padding.Empty };
            Controls.Remove(canvas);
            canvas.Dock = DockStyle.None;
            canvas.Location = Point.Empty;
            canvas.Margin = Padding.Empty;
            viewport.Controls.Add(canvas);
            Controls.Add(viewport);
            Controls.SetChildIndex(viewport, index);
            return viewport;
        }

        private void LayoutViewports()
        {
            if (_layingOutViewports || _contentViewport == null || _actionViewport == null) return;
            _layingOutViewports = true;
            try
            {
                bool footerScroll = _actionDesignSize.Height > 0 && _actionViewport.Width < Px(_actionMinimumWidth);
                _actionViewport.Height = Px(_actionDesignSize.Height) +
                    (footerScroll ? SystemInformation.HorizontalScrollBarHeight : 0);
                ArrangeViewport(_actionViewport, ActionBar,
                    new Size(Px(_actionMinimumWidth), Px(_actionDesignSize.Height)), false);
                ArrangeViewport(_contentViewport, ContentPanel,
                    ReflowContent ? Size.Empty : new Size(Px(_contentDesignSize.Width), Px(_contentDesignSize.Height)));
            }
            finally { _layingOutViewports = false; }
        }

        private static void ArrangeViewport(Panel viewport, Panel canvas, Size minimum, bool fillHeight = true)
        {
            var scroll = viewport.AutoScrollPosition;
            viewport.SuspendLayout();
            try
            {
                // Drop the cached DisplayRectangle from the previous DPI/size.
                // WinForms otherwise retains it for anchored children after shrinking.
                viewport.AutoScroll = false;
                viewport.AutoScrollMinSize = Size.Empty;
                canvas.Location = Point.Empty;
                bool overflow = minimum.Width > viewport.ClientSize.Width || minimum.Height > viewport.ClientSize.Height;
                if (overflow)
                {
                    viewport.AutoScrollMinSize = minimum;
                    viewport.AutoScroll = true;
                }
                canvas.Size = new Size(Math.Max(minimum.Width, viewport.ClientSize.Width),
                    fillHeight ? Math.Max(minimum.Height, viewport.ClientSize.Height) : minimum.Height);
            }
            finally { viewport.ResumeLayout(true); }
            // Account for the space consumed by scrollbars after layout.
            canvas.Size = new Size(Math.Max(minimum.Width, viewport.ClientSize.Width),
                fillHeight ? Math.Max(minimum.Height, viewport.ClientSize.Height) : minimum.Height);
            if (viewport.AutoScroll)
                viewport.AutoScrollPosition = new Point(-scroll.X, -scroll.Y);
        }

        protected override void OnLayoutDpiChanged()
        {
            base.OnLayoutDpiChanged();
            LayoutViewports();
        }

        // --- Properties ---
        public string Title { get; set; } = "Ўзбек Орфо";
        public Image TitleIcon { get; set; }
        public bool ShowMinimizeButton { get; set; } = false;

        /// <summary>
        /// Whether the form can be resized by dragging edges. Default true.
        /// </summary>
        public bool AllowResize { get; set; } = true;

        /// <summary>
        /// The content panel below the title bar — add your controls here.
        /// </summary>
        public Panel ContentPanel { get; private set; }

        /// <summary>
        /// The bottom action bar — add OK/Cancel buttons here.
        /// </summary>
        public Panel ActionBar { get; private set; }

        public ModernForm()
        {
            // Base form settings
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw | ControlStyles.OptimizedDoubleBuffer
                     | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
            BackColor = ThemeManager.Background;
            ForeColor = ThemeManager.TextPrimary;
            Font = ThemeManager.FontBase;
            MinimumSize = new Size(320, 200);
            Padding = new Padding(RESIZE_BORDER, 1, RESIZE_BORDER, RESIZE_BORDER);

            SetupLayout();
        }

        private void SetupLayout()
        {
            // Content panel — fills the area between title bar and action bar
            ContentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeManager.Background,
                Padding = new Padding(ThemeManager.SpaceXL, Px(ThemeManager.SpaceLG),
                                      ThemeManager.SpaceXL, Px(ThemeManager.SpaceLG)),
                AutoScroll = false
            };

            // Action bar at bottom
            ActionBar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 56,
                BackColor = ThemeManager.Surface,
                Padding = new Padding(ThemeManager.SpaceXL, ThemeManager.SpaceMD,
                                      ThemeManager.SpaceXL, ThemeManager.SpaceMD)
            };

            // Top spacer for custom title bar — forwards mouse events to form
            _titleSpacer = new Panel
            {
                Dock = DockStyle.Top,
                Height = TITLE_BAR_HEIGHT,
                BackColor = Color.Transparent
            };
            _titleSpacer.MouseDown += TitleSpacer_MouseDown;
            _titleSpacer.MouseMove += TitleSpacer_MouseMove;
            _titleSpacer.MouseUp += TitleSpacer_MouseUp;
            _titleSpacer.MouseLeave += TitleSpacer_MouseLeave;

            // Separator line above action bar
            var separator = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 1,
                BackColor = ThemeManager.Border
            };

            Controls.Add(ContentPanel);
            Controls.Add(separator);
            Controls.Add(ActionBar);
            Controls.Add(_titleSpacer);
        }

        // =====================================================================
        //  ROUNDED CORNERS
        // =====================================================================

        private const int WM_ENTERSIZEMOVE = 0x0231;
        private const int WM_EXITSIZEMOVE  = 0x0232;
        private bool _isResizing;

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            // Skip expensive Region recalc during live drag — apply on release
            if (!_isResizing)
                ApplyRoundedCorners();
        }

        private void ApplyRoundedCorners()
        {
            int radius = Px(ThemeManager.RadiusLG);
            using (var path = CreateRoundedRectPath(new Rectangle(0, 0, Width, Height), radius))
            {
                var previousRegion = Region;
                Region = new Region(path);
                previousRegion?.Dispose();
            }
        }

        private static GraphicsPath CreateRoundedRectPath(Rectangle rect, int radius)
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

        // =====================================================================
        //  CUSTOM PAINTING — Title bar, border, close/min buttons
        // =====================================================================

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Border
            using (var pen = new Pen(ThemeManager.Border, 1))
            using (var path = CreateRoundedRectPath(new Rectangle(0, 0, Width - 1, Height - 1), Px(ThemeManager.RadiusLG)))
            {
                g.DrawPath(pen, path);
            }

            // Paint bottom resize grip area to blend with ActionBar
            if (AllowResize)
            {
                using (var brush = new SolidBrush(ThemeManager.Surface))
                {
                    g.FillRectangle(brush, 0, Height - RESIZE_BORDER, Width, RESIZE_BORDER);
                }
            }

            // Title bar background
            var titleRect = new Rectangle(0, 0, Width, TITLE_BAR_HEIGHT);
            using (var brush = new SolidBrush(ThemeManager.Surface))
            {
                g.FillRectangle(brush, titleRect);
            }

            // Title bar bottom border
            using (var pen = new Pen(ThemeManager.Border))
            {
                g.DrawLine(pen, 0, TITLE_BAR_HEIGHT, Width, TITLE_BAR_HEIGHT);
            }

            // Title icon + text — vertically centered
            int iconSize = Px(24);
            int iconY = (TITLE_BAR_HEIGHT - iconSize) / 2;
            int textX = Px(ThemeManager.SpaceLG);

            if (TitleIcon != null)
            {
                g.DrawImage(TitleIcon, new Rectangle(Px(ThemeManager.SpaceLG), iconY, iconSize, iconSize));
                textX = Px(ThemeManager.SpaceLG) + iconSize + Px(ThemeManager.SpaceSM);
            }

            // Title text — vertically centered, bigger font
            using (var brush = new SolidBrush(ThemeManager.TextPrimary))
            {
                var titleFont = UiFont(ThemeManager.FontLGBold);
                int buttonsWidth = Px(ShowMinimizeButton ? 88 : 48);
                using (var format = new StringFormat { LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
                    g.DrawString(Title, titleFont, brush,
                        new RectangleF(textX, 0, Math.Max(1, Width - textX - buttonsWidth), TITLE_BAR_HEIGHT), format);
            }

            // Close button
            int btnSize = Px(32);
            int btnY = (TITLE_BAR_HEIGHT - btnSize) / 2;
            _closeBtnRect = new Rectangle(Width - btnSize - Px(8), btnY, btnSize, btnSize);
            DrawTitleButton(g, _closeBtnRect, "✕", _closeHovered,
                _closeHovered ? ThemeManager.ButtonDanger : Color.Transparent,
                _closeHovered ? Color.White : ThemeManager.TextSecondary);

            // Minimize button
            if (ShowMinimizeButton)
            {
                _minBtnRect = new Rectangle(Width - btnSize * 2 - Px(12), btnY, btnSize, btnSize);
                DrawTitleButton(g, _minBtnRect, "─", _minHovered,
                    _minHovered ? ThemeManager.SurfaceHover : Color.Transparent,
                    ThemeManager.TextSecondary);
            }
        }

        private void DrawTitleButton(Graphics g, Rectangle rect, string symbol,
            bool hovered, Color bg, Color fg)
        {
            if (hovered)
            {
                using (var brush = new SolidBrush(bg))
                using (var path = CreateRoundedRectPath(rect, Px(ThemeManager.RadiusSM)))
                {
                    g.FillPath(brush, path);
                }
            }

            using (var brush = new SolidBrush(fg))
            {
                var sf = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                g.DrawString(symbol, UiFont(ThemeManager.FontBase), brush, rect, sf);
            }
        }

        // =====================================================================
        //  MOUSE HANDLING — Drag, close, minimize
        // =====================================================================

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);

            if (_closeBtnRect.Contains(e.Location))
            {
                Close();
                return;
            }

            if (ShowMinimizeButton && _minBtnRect.Contains(e.Location))
            {
                WindowState = FormWindowState.Minimized;
                return;
            }

            // Drag window by title bar
            if (e.Y <= TITLE_BAR_HEIGHT && e.Button == MouseButtons.Left)
            {
                _isDragging = true;
                _dragStart = e.Location;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);

            if (_isDragging)
            {
                Location = new Point(
                    Location.X + e.X - _dragStart.X,
                    Location.Y + e.Y - _dragStart.Y);
                return;
            }

            // Hover tracking for title buttons
            bool newCloseHover = _closeBtnRect.Contains(e.Location);
            bool newMinHover = ShowMinimizeButton && _minBtnRect.Contains(e.Location);

            if (newCloseHover != _closeHovered || newMinHover != _minHovered)
            {
                _closeHovered = newCloseHover;
                _minHovered = newMinHover;
                Invalidate(new Rectangle(0, 0, Width, TITLE_BAR_HEIGHT));
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            bool wasDragging = _isDragging;
            _isDragging = false;
            if (wasDragging) FitToScreen();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (_closeHovered || _minHovered)
            {
                _closeHovered = false;
                _minHovered = false;
                Invalidate(new Rectangle(0, 0, Width, TITLE_BAR_HEIGHT));
            }
        }

        // =====================================================================
        //  TITLE SPACER EVENT FORWARDING
        //  The titleSpacer panel covers the title bar area and intercepts
        //  mouse events. We convert coordinates and forward to the form.
        // =====================================================================

        private void TitleSpacer_MouseDown(object sender, MouseEventArgs e)
        {
            var formPt = PointToClient(_titleSpacer.PointToScreen(e.Location));
            OnMouseDown(new MouseEventArgs(e.Button, e.Clicks, formPt.X, formPt.Y, e.Delta));
        }

        private void TitleSpacer_MouseMove(object sender, MouseEventArgs e)
        {
            var formPt = PointToClient(_titleSpacer.PointToScreen(e.Location));
            OnMouseMove(new MouseEventArgs(e.Button, e.Clicks, formPt.X, formPt.Y, e.Delta));
        }

        private void TitleSpacer_MouseUp(object sender, MouseEventArgs e)
        {
            var formPt = PointToClient(_titleSpacer.PointToScreen(e.Location));
            OnMouseUp(new MouseEventArgs(e.Button, e.Clicks, formPt.X, formPt.Y, e.Delta));
        }

        private void TitleSpacer_MouseLeave(object sender, EventArgs e)
        {
            OnMouseLeave(e);
        }

        // =====================================================================
        //  KEYBOARD — Escape to close
        // =====================================================================

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // =====================================================================
        //  RESIZE SUPPORT — WM_NCHITTEST for edge dragging
        // =====================================================================

        private const int WM_NCHITTEST = 0x0084;
        private const int HTCLIENT = 1;
        private const int HTLEFT = 10;
        private const int HTRIGHT = 11;
        private const int HTTOP = 12;
        private const int HTTOPLEFT = 13;
        private const int HTTOPRIGHT = 14;
        private const int HTBOTTOM = 15;
        private const int HTBOTTOMLEFT = 16;
        private const int HTBOTTOMRIGHT = 17;

        protected override void WndProc(ref Message m)
        {
            // --- Live resize: suspend / resume layout ---
            if (m.Msg == WM_ENTERSIZEMOVE)
            {
                _isResizing = true;
                SuspendLayout();
                // Temporarily remove region for smooth resize
                var previousRegion = Region;
                Region = null;
                previousRegion?.Dispose();
                base.WndProc(ref m);
                return;
            }
            if (m.Msg == WM_EXITSIZEMOVE)
            {
                _isResizing = false;
                ResumeLayout(true);
                ApplyRoundedCorners();
                Invalidate(true);
                base.WndProc(ref m);
                return;
            }

            if (AllowResize && m.Msg == WM_NCHITTEST)
            {
                base.WndProc(ref m);

                // Screen coords from LPARAM
                int x = (short)(m.LParam.ToInt32() & 0xFFFF);
                int y = (short)((m.LParam.ToInt32() >> 16) & 0xFFFF);
                var pt = PointToClient(new Point(x, y));

                bool left = pt.X < RESIZE_BORDER;
                bool right = pt.X >= Width - RESIZE_BORDER;
                bool top = pt.Y < RESIZE_BORDER;
                bool bottom = pt.Y >= Height - RESIZE_BORDER;

                if (top && left) m.Result = (IntPtr)HTTOPLEFT;
                else if (top && right) m.Result = (IntPtr)HTTOPRIGHT;
                else if (bottom && left) m.Result = (IntPtr)HTBOTTOMLEFT;
                else if (bottom && right) m.Result = (IntPtr)HTBOTTOMRIGHT;
                else if (left) m.Result = (IntPtr)HTLEFT;
                else if (right) m.Result = (IntPtr)HTRIGHT;
                else if (top) m.Result = (IntPtr)HTTOP;
                else if (bottom) m.Result = (IntPtr)HTBOTTOM;

                return;
            }

            base.WndProc(ref m);
        }

        // =====================================================================
        //  DROP SHADOW (Windows API)
        // =====================================================================

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ClassStyle |= 0x00020000; // CS_DROPSHADOW
                // WS_EX_COMPOSITED — flicker-free child control painting
                cp.ExStyle |= 0x02000000;
                return cp;
            }
        }
    }
}
