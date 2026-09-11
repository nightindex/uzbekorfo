using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.UI.Controls
{
    /// <summary>
    /// A thin, modern, themed vertical scrollbar control.
    /// Designed to overlay alongside a DataGridView or Panel, replacing
    /// the default ugly Windows scrollbar.
    /// 
    /// Usage:
    ///   var scrollBar = ModernScrollBar.AttachTo(dataGridView);
    /// </summary>
    public class ModernScrollBar : Control
    {
        private int Px(int value) => DpiLayout.Pixels(this, value);
        private Font UiFont(Font value) => DpiLayout.Font(this, value);

        // =====================================================================
        //  CONSTANTS
        // =====================================================================

        private const int DEFAULT_BAR_WIDTH = 10;   // Total control width
        private const int DEFAULT_THUMB_WIDTH = 6;   // Visible thumb width (centered)
        private int THUMB_MIN_HEIGHT => Px(30);    // Minimum thumb size
        private int THUMB_RADIUS => Px(3);         // Rounded corners
        private int TRACK_PADDING => Px(2);        // Top/bottom padding

        private int _barWidth = DEFAULT_BAR_WIDTH;
        private int _thumbWidth = DEFAULT_THUMB_WIDTH;

        /// <summary>Total scrollbar control width in pixels (default 10).</summary>
        public int BarWidth
        {
            get => _barWidth;
            set
            {
                _barWidth = Math.Max(4, value);
                int maxThumb = Math.Max(2, _barWidth - 1);
                if (_thumbWidth > maxThumb) _thumbWidth = maxThumb;
                Width = Px(_barWidth);
                UpdateThumb();
            }
        }

        /// <summary>Visible thumb width in pixels, centered within BarWidth (default 6).</summary>
        public int ThumbWidth
        {
            get => _thumbWidth;
            set
            {
                int maxThumb = Math.Max(2, _barWidth - 1);
                _thumbWidth = Math.Max(2, Math.Min(value, maxThumb));
                UpdateThumb();
            }
        }
        private const int LVM_FIRST = 0x1000;
        private const int LVM_SCROLL = LVM_FIRST + 20;
        private const int LVM_GETTOPINDEX = LVM_FIRST + 39;
        private const int SB_VERT = 1;

        // =====================================================================
        //  STATE
        // =====================================================================

        private int _maximum = 100;
        private int _viewportSize = 10;
        private int _value;

        private bool _thumbHovered;
        private bool _thumbPressed;
        private int _dragStartY;
        private int _dragStartValue;
        private Rectangle _thumbRect;

        // --- Fade animation: scrollbar appears on scroll, fades after idle ---
        private Timer _fadeTimer;
        private float _opacity = 0f;
        private bool _showing;

        // =====================================================================
        //  PROPERTIES
        // =====================================================================

        public int Maximum
        {
            get => _maximum;
            set { _maximum = Math.Max(1, value); UpdateThumb(); }
        }

        public int ViewportSize
        {
            get => _viewportSize;
            set { _viewportSize = Math.Max(1, value); UpdateThumb(); }
        }

        public int Value
        {
            get => _value;
            set
            {
                int clamped = Clamp(value, 0, Math.Max(0, _maximum - _viewportSize));
                if (clamped != _value)
                {
                    _value = clamped;
                    UpdateThumb();
                    ValueChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        public event EventHandler ValueChanged;

        // =====================================================================
        //  CONSTRUCTOR
        // =====================================================================

        public ModernScrollBar()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);

            Width = Px(_barWidth);
            Cursor = Cursors.Default;
            BackColor = ThemeManager.Background;

            _fadeTimer = new Timer { Interval = 30 };
            _fadeTimer.Tick += FadeTimer_Tick;
        }

        // =====================================================================
        //  SHOW / HIDE with animation
        // =====================================================================

        public void ShowScrollBar()
        {
            _showing = true;
            _opacity = 1f;
            Visible = true;
            Invalidate();
            _fadeTimer.Stop();
            _fadeTimer.Start();
        }

        private void FadeTimer_Tick(object sender, EventArgs e)
        {
            if (_thumbPressed) return; // don't fade while dragging

            if (_showing)
            {
                // Stay visible for a moment then begin fading
                _showing = false;
                _fadeTimer.Interval = 1200; // wait before fade
                return;
            }

            _fadeTimer.Interval = 30;
            _opacity -= 0.08f;
            if (_opacity <= 0f)
            {
                _opacity = 0f;
                _fadeTimer.Stop();
            }
            Invalidate();
        }

        // =====================================================================
        //  PAINTING
        // =====================================================================

        protected override void OnPaint(PaintEventArgs e)
        {
            if (_maximum <= _viewportSize) return; // no scroll needed

            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int alpha = (int)(255 * _opacity);
            if (alpha <= 0) return;

            // Track background (subtle)
            using (var trackBrush = new SolidBrush(Color.FromArgb(alpha / 3, ThemeManager.ScrollTrack)))
            {
                using (var trackPath = CreateRoundedRect(
                    new Rectangle((Width - Px(_thumbWidth)) / 2, TRACK_PADDING,
                                  Px(_thumbWidth), Math.Max(1, Height - TRACK_PADDING * 2)),
                    THUMB_RADIUS))
                {
                    g.FillPath(trackBrush, trackPath);
                }
            }

            // Thumb
            Color thumbColor = _thumbPressed ? ThemeManager.ScrollThumbPressed
                             : _thumbHovered ? ThemeManager.ScrollThumbHover
                             : ThemeManager.ScrollThumb;

            using (var thumbBrush = new SolidBrush(Color.FromArgb(alpha, thumbColor)))
            using (var thumbPath = CreateRoundedRect(_thumbRect, THUMB_RADIUS))
            {
                g.FillPath(thumbBrush, thumbPath);
            }
        }

        // =====================================================================
        //  THUMB CALCULATION
        // =====================================================================

        private void UpdateThumb()
        {
            if (_maximum <= _viewportSize)
            {
                _thumbRect = Rectangle.Empty;
                Invalidate();
                return;
            }

            int trackHeight = Math.Max(1, Height - TRACK_PADDING * 2);
            float ratio = (float)_viewportSize / _maximum;
            int thumbH = Math.Min(trackHeight, Math.Max(THUMB_MIN_HEIGHT, (int)(trackHeight * ratio)));

            int scrollRange = _maximum - _viewportSize;
            float scrollRatio = scrollRange > 0 ? (float)_value / scrollRange : 0;
            int thumbY = TRACK_PADDING + (int)((trackHeight - thumbH) * scrollRatio);

            int thumbX = (Width - Px(_thumbWidth)) / 2;
            _thumbRect = new Rectangle(thumbX, thumbY, Px(_thumbWidth), thumbH);
            Invalidate();
        }

        // =====================================================================
        //  MOUSE HANDLING
        // =====================================================================

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (_thumbRect.Contains(e.Location))
            {
                _thumbPressed = true;
                _dragStartY = e.Y;
                _dragStartValue = _value;
                Capture = true;
                Invalidate();
            }
            else if (e.Y < _thumbRect.Y)
            {
                // Click above thumb — page up
                Value -= _viewportSize;
            }
            else if (e.Y > _thumbRect.Bottom)
            {
                // Click below thumb — page down
                Value += _viewportSize;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_thumbPressed)
            {
                int dy = e.Y - _dragStartY;
                int trackHeight = Height - TRACK_PADDING * 2 - _thumbRect.Height;
                int scrollRange = _maximum - _viewportSize;
                if (trackHeight > 0 && scrollRange > 0)
                {
                    int newVal = _dragStartValue + (int)((float)dy / trackHeight * scrollRange);
                    Value = newVal;
                }
            }
            else
            {
                bool wasHovered = _thumbHovered;
                _thumbHovered = _thumbRect.Contains(e.Location);
                if (wasHovered != _thumbHovered) Invalidate();
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (_thumbPressed)
            {
                _thumbPressed = false;
                Capture = false;
                Invalidate();
                ShowScrollBar(); // restart fade timer
            }
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            _opacity = 1f;
            _fadeTimer.Stop();
            Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _thumbHovered = false;
            ShowScrollBar(); // begin fade
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UpdateThumb();
        }

        // =====================================================================
        //  HELPER
        // =====================================================================

        private static GraphicsPath CreateRoundedRect(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = Math.Max(1, Math.Min(radius * 2, Math.Min(rect.Width, rect.Height)));
            if (rect.Width <= 0 || rect.Height <= 0) { path.AddRectangle(rect); return path; }

            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private static int Clamp(int val, int min, int max)
        {
            if (val < min) return min;
            if (val > max) return max;
            return val;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ShowScrollBar(IntPtr hWnd, int wBar, bool bShow);

        private static int GetListViewTopIndex(ListView listView)
        {
            if (listView == null || !listView.IsHandleCreated) return 0;
            return (int)SendMessage(listView.Handle, LVM_GETTOPINDEX, IntPtr.Zero, IntPtr.Zero);
        }

        // =====================================================================
        //  STATIC FACTORY — Attach to a DataGridView
        // =====================================================================

        /// <summary>
        /// Creates a modern scrollbar and attaches it to a DataGridView.
        /// Hides the native vertical scrollbar and syncs scroll position.
        /// </summary>
        public static ModernScrollBar AttachTo(DataGridView grid)
        {
            if (grid == null) throw new ArgumentNullException(nameof(grid));

            // Keep native scrollbar functional (needed for mouse wheel + keyboard)
            // but we overlay our custom scrollbar on top
            grid.ScrollBars = ScrollBars.Both;

            var sb = new ModernScrollBar
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            sb.BarWidth = DEFAULT_BAR_WIDTH;

            // Position the scrollbar inside the grid's parent
            Action positionBar = () =>
            {
                sb.Location = new Point(grid.Right - sb.Width, grid.Top);
                sb.Height = grid.Height;
                sb.BringToFront();
            };

            // Sync scrollbar metrics from grid
            Action syncMetrics = () =>
            {
                int totalRows = grid.RowCount;
                int visibleRows = grid.DisplayedRowCount(true);
                sb.Maximum = totalRows;
                sb.ViewportSize = visibleRows;

                int first = grid.FirstDisplayedScrollingRowIndex;
                if (first >= 0)
                    sb._value = first; // set without triggering event
                sb.UpdateThumb();
            };

            // When our scrollbar changes, scroll the grid
            sb.ValueChanged += (s, e) =>
            {
                if (grid.RowCount > 0 && sb.Value >= 0 && sb.Value < grid.RowCount)
                {
                    grid.FirstDisplayedScrollingRowIndex = sb.Value;
                }
            };

            // When grid scrolls (keyboard, mouse wheel), update our scrollbar
            grid.Scroll += (s, e) =>
            {
                if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
                {
                    int first = grid.FirstDisplayedScrollingRowIndex;
                    if (first >= 0)
                    {
                        sb._value = first;
                        sb.UpdateThumb();
                        sb.ShowScrollBar();
                    }
                }
            };

            // Mouse wheel on grid — sync our scrollbar after native scroll
            grid.MouseWheel += (s, e) =>
            {
                // Small delay so native scroll finishes first
                var timer = new Timer { Interval = 20 };
                timer.Tick += (t, te) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    syncMetrics();
                    sb.ShowScrollBar();
                };
                timer.Start();
            };

            grid.RowsAdded += (s, e) => syncMetrics();
            grid.RowsRemoved += (s, e) => syncMetrics();

            grid.Resize += (s, e) =>
            {
                positionBar();
                syncMetrics();
            };

            // Add to parent when available
            if (grid.Parent != null)
            {
                grid.Parent.Controls.Add(sb);
                positionBar();
                syncMetrics();
            }
            else
            {
                grid.ParentChanged += (s, e) =>
                {
                    if (grid.Parent != null)
                    {
                        grid.Parent.Controls.Add(sb);
                        positionBar();
                        syncMetrics();
                    }
                };
            }

            return sb;
        }

        /// <summary>
        /// Creates a modern scrollbar and attaches it to a ListView (including VirtualMode).
        /// The native scrollbar remains functional underneath; this control overlays it.
        /// </summary>
        public static ModernScrollBar AttachTo(ListView listView)
        {
            if (listView == null) throw new ArgumentNullException(nameof(listView));

            var sb = new ModernScrollBar
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            sb.BarWidth = Math.Max(SystemInformation.VerticalScrollBarWidth, DEFAULT_BAR_WIDTH);

            int lastTop = -1;
            int lastCount = -1;

            Func<int> getTotalRows = () =>
                listView.VirtualMode ? Math.Max(0, listView.VirtualListSize) : listView.Items.Count;

            Func<int> getRowHeight = () =>
            {
                int fromImages = listView.SmallImageList != null
                    ? listView.SmallImageList.ImageSize.Height
                    : 0;
                return Math.Max(22, fromImages > 0 ? fromImages : listView.Font.Height + 14);
            };

            Action positionBar = () =>
            {
                if (listView.Parent == null) return;
                sb.Location = new Point(listView.Right - sb.Width, listView.Top);
                sb.Height = listView.Height;
                sb.BringToFront();
            };

            Action syncMetrics = () =>
            {
                if (listView.IsDisposed || !listView.IsHandleCreated) return;

                int totalRows = getTotalRows();
                int rowHeight = getRowHeight();
                int visibleRows = Math.Max(1, listView.ClientSize.Height / Math.Max(1, rowHeight));

                sb.Maximum = Math.Max(1, totalRows);
                sb.ViewportSize = Math.Max(1, Math.Min(totalRows > 0 ? totalRows : 1, visibleRows));

                int top = GetListViewTopIndex(listView);
                if (top < 0) top = 0;

                sb._value = Clamp(top, 0, Math.Max(0, sb._maximum - sb._viewportSize));
                sb.UpdateThumb();

                bool needsScroll = totalRows > visibleRows;
                sb.Visible = needsScroll;
                if (needsScroll && (top != lastTop || totalRows != lastCount))
                    sb.ShowScrollBar();

                lastTop = top;
                lastCount = totalRows;
            };

            sb.ValueChanged += (s, e) =>
            {
                if (listView.IsDisposed || !listView.IsHandleCreated) return;

                int totalRows = getTotalRows();
                if (totalRows <= 0) return;

                int target = Clamp(sb.Value, 0, totalRows - 1);
                int currentTop = GetListViewTopIndex(listView);
                if (currentTop < 0) currentTop = 0;
                if (target == currentTop) return;

                int rowHeight = getRowHeight();
                int deltaRows = target - currentTop;

                // Pixel-scroll first so dragging the custom thumb feels smooth.
                if (Math.Abs(deltaRows) <= 300)
                {
                    int pixelDelta = deltaRows * Math.Max(1, rowHeight);
                    SendMessage(listView.Handle, LVM_SCROLL, IntPtr.Zero, (IntPtr)pixelDelta);
                }
                else
                {
                    try { listView.EnsureVisible(target); } catch { }
                }

                syncMetrics();
            };

            Action queueSync = () =>
            {
                if (listView.IsDisposed) return;
                try
                {
                    listView.BeginInvoke((Action)(() =>
                    {
                        if (listView.IsDisposed) return;
                        positionBar();
                        syncMetrics();
                    }));
                }
                catch { }
            };

            listView.Resize += (s, e) => queueSync();
            listView.HandleCreated += (s, e) => queueSync();
            listView.MouseWheel += (s, e) => queueSync();
            listView.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down ||
                    e.KeyCode == Keys.PageUp || e.KeyCode == Keys.PageDown ||
                    e.KeyCode == Keys.Home || e.KeyCode == Keys.End)
                {
                    queueSync();
                }
            };

            var poll = new Timer { Interval = 120 };
            poll.Tick += (s, e) =>
            {
                if (listView.IsDisposed || sb.IsDisposed)
                {
                    poll.Stop();
                    poll.Dispose();
                    return;
                }

                syncMetrics();
            };
            poll.Start();

            listView.Disposed += (s, e) =>
            {
                try { poll.Stop(); poll.Dispose(); } catch { }
                try { if (!sb.IsDisposed) sb.Dispose(); } catch { }
            };

            if (listView.Parent != null)
            {
                listView.Parent.Controls.Add(sb);
                positionBar();
                syncMetrics();
            }
            else
            {
                listView.ParentChanged += (s, e) =>
                {
                    if (listView.Parent != null)
                    {
                        listView.Parent.Controls.Add(sb);
                        positionBar();
                        syncMetrics();
                    }
                };
            }

            return sb;
        }

        /// <summary>
        /// Creates a modern scrollbar and attaches it to a scrollable Panel.
        /// </summary>
        public static ModernScrollBar AttachTo(Panel panel, int totalWidth = DEFAULT_BAR_WIDTH)
        {
            if (panel == null) throw new ArgumentNullException(nameof(panel));

            int effectiveWidth = Math.Max(DEFAULT_THUMB_WIDTH + 1, totalWidth);
            var sb = new ModernScrollBar
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            sb.BarWidth = effectiveWidth;

            Action positionBar = () =>
            {
                if (panel.Parent == null) return;
                sb.Location = new Point(panel.Right - sb.Width, panel.Top);
                sb.Height = panel.Height;
                sb.BringToFront();
            };

            Action syncMetrics = () =>
            {
                if (panel.IsDisposed || sb.IsDisposed) return;

                // Hide native vertical scrollbar so only the modern overlay is visible.
                if (panel.IsHandleCreated)
                {
                    try { ShowScrollBar(panel.Handle, SB_VERT, false); } catch { }
                }

                int maxScroll = panel.VerticalScroll.Maximum - panel.ClientSize.Height;
                if (maxScroll <= 0) { sb.Visible = false; return; }
                sb.Maximum = panel.VerticalScroll.Maximum;
                sb.ViewportSize = panel.ClientSize.Height;
                sb._value = panel.VerticalScroll.Value;
                sb.UpdateThumb();
                sb.Visible = true;
            };

            sb.ValueChanged += (s, e) =>
            {
                if (panel.IsDisposed) return;
                panel.AutoScrollPosition = new Point(0, sb.Value);
            };

            panel.Scroll += (s, e) =>
            {
                if (panel.IsDisposed || sb.IsDisposed) return;
                sb._value = panel.VerticalScroll.Value;
                sb.UpdateThumb();
                sb.ShowScrollBar();
                if (panel.IsHandleCreated)
                {
                    try { ShowScrollBar(panel.Handle, SB_VERT, false); } catch { }
                }
            };

            panel.Resize += (s, e) =>
            {
                positionBar();
                syncMetrics();
            };
            panel.LocationChanged += (s, e) => positionBar();
            panel.ControlAdded += (s, e) => syncMetrics();
            panel.ControlRemoved += (s, e) => syncMetrics();
            panel.VisibleChanged += (s, e) => syncMetrics();
            panel.ParentChanged += (s, e) =>
            {
                if (panel.Parent == null) return;
                if (sb.Parent != panel.Parent)
                {
                    if (sb.Parent != null) sb.Parent.Controls.Remove(sb);
                    panel.Parent.Controls.Add(sb);
                }
                positionBar();
                syncMetrics();
            };
            panel.Disposed += (s, e) =>
            {
                try { if (!sb.IsDisposed) sb.Dispose(); } catch { }
            };
            panel.HandleCreated += (s, e) =>
            {
                try { ShowScrollBar(panel.Handle, SB_VERT, false); } catch { }
                syncMetrics();
            };

            if (panel.Parent != null)
            {
                panel.Parent.Controls.Add(sb);
                positionBar();
                syncMetrics();
            }

            return sb;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _fadeTimer?.Stop();
                _fadeTimer?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
