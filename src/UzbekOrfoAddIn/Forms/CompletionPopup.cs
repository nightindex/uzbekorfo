using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>Compact, non-activating alternatives popup owned by Word.</summary>
    public sealed class CompletionPopup : Form
    {
        private string[] _items = new string[0];
        private Font _ownedFont;
        private int _dpi = 96;
        private Color _surface;
        private Color _text;
        private Color _secondaryText;
        private Color _selectedSurface;
        private Color _border;

        public int SelectedIndex { get; private set; }
        private int RowHeight => ScreenGeometry.Scale(38, _dpi);
        public event Action<string> Accepted;
        public event Action Dismissed;
        public string Selected => _items.Length == 0 ? null : _items[SelectedIndex];
        internal Rectangle PresentationBounds { get; private set; }
        internal int PresentationDpi => _dpi;

        protected override bool ShowWithoutActivation => true;
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x08000000 | 0x80; // WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW
                cp.ClassStyle |= 0x00020000;      // CS_DROPSHADOW
                return cp;
            }
        }

        public CompletionPopup()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
            DoubleBuffered = true;
            AccessibleName = "MatnAI сўз таклифлари";
            SetAppearance(96, false);
        }

        public void Present(string[] items, Rectangle caret, IWin32Window owner)
        {
            IntPtr ownerHandle = owner == null ? IntPtr.Zero : owner.Handle;
            Present(items, OverlayPositioner.CreateAnchor(caret, ownerHandle,
                "Segoe UI", 10f, FontStyle.Regular), owner);
        }

        internal void Present(string[] items, OverlayAnchor anchor, IWin32Window owner)
        {
            if (items == null || items.Length == 0 || anchor == null)
            {
                Hide();
                return;
            }

            bool changed = !SameItems(items, _items);
            _items = (string[])items.Clone();
            if (changed || SelectedIndex >= _items.Length) SelectedIndex = 0;
            SetAppearance(anchor.Dpi, anchor.DarkBackground);

            int horizontalPadding = ScreenGeometry.Scale(28, _dpi);
            int measured = 0;
            foreach (string item in _items)
                measured = Math.Max(measured, TextRenderer.MeasureText(item ?? string.Empty, Font,
                    Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Width);
            int minWidth = ScreenGeometry.Scale(236, _dpi);
            int maxWidth = ScreenGeometry.Scale(360, _dpi);
            var desired = new Size(Math.Max(minWidth, Math.Min(maxWidth, measured + horizontalPadding)),
                _items.Length * RowHeight + 2);
            Rectangle bounds = OverlayPositioner.PlacePopup(anchor, desired);
            if (bounds.IsEmpty)
            {
                Hide();
                return;
            }

            using (new DpiLayout.Context(DpiLayout.WindowContext(anchor.OwnerHandle)))
            {
                if (!IsHandleCreated) { var unused = Handle; }
                if (!Visible)
                {
                    if (owner == null) Show(); else Show(owner);
                }
            }
            OverlayPositioner.MovePhysical(this, bounds, anchor.OwnerHandle);
            PresentationBounds = bounds;
            AccessibleDescription = string.Join(", ", _items);
            AccessibilityNotifyClients(AccessibleEvents.Reorder, -1);
            UpdateRoundedRegion();
            Invalidate();
        }

        public void MoveSelection(int direction)
        {
            if (_items.Length == 0 || !Visible) return;
            SelectedIndex = (SelectedIndex + direction + _items.Length) % _items.Length;
            AccessibilityNotifyClients(AccessibleEvents.Selection, SelectedIndex);
            Invalidate();
        }

        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0x21) // WM_MOUSEACTIVATE
            {
                message.Result = new IntPtr(3); // MA_NOACTIVATE
                return;
            }
            base.WndProc(ref message);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.Clear(_surface);
            for (int index = 0; index < _items.Length; index++)
            {
                var row = new Rectangle(1, index * RowHeight + 1, Width - 2, RowHeight);
                if (index == SelectedIndex)
                {
                    using (var selected = new SolidBrush(_selectedSurface)) e.Graphics.FillRectangle(selected, row);
                    using (var accent = new SolidBrush(ThemeManager.Primary))
                        e.Graphics.FillRectangle(accent, row.Left, row.Top, ScreenGeometry.Scale(3, _dpi), row.Height);
                }
                TextRenderer.DrawText(e.Graphics, _items[index], Font,
                    new Rectangle(row.Left + ScreenGeometry.Scale(13, _dpi), row.Top,
                        row.Width - ScreenGeometry.Scale(22, _dpi), row.Height),
                    index == SelectedIndex ? _text : _secondaryText,
                    TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix |
                    TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter);
            }
            using (var pen = new Pen(_border))
            using (var path = RoundedPath(new Rectangle(0, 0, Width - 1, Height - 1),
                ScreenGeometry.Scale(9, _dpi)))
                e.Graphics.DrawPath(pen, path);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int index = e.Y / Math.Max(1, RowHeight);
            if (index >= 0 && index < _items.Length && index != SelectedIndex)
            {
                SelectedIndex = index;
                AccessibilityNotifyClients(AccessibleEvents.Selection, index);
                Invalidate();
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Right)
            {
                Dismissed?.Invoke();
                return;
            }
            if (e.Button != MouseButtons.Left) return;
            int index = e.Y / Math.Max(1, RowHeight);
            if (index >= 0 && index < _items.Length) Accepted?.Invoke(_items[index]);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UpdateRoundedRegion();
        }

        private void SetAppearance(int dpi, bool dark)
        {
            dpi = Math.Max(96, dpi);
            if (_dpi != dpi || _ownedFont == null)
            {
                _dpi = dpi;
                var replacement = new Font("Segoe UI", 10f * dpi / 72f,
                    FontStyle.Regular, GraphicsUnit.Pixel);
                var previous = _ownedFont;
                _ownedFont = replacement;
                Font = replacement;
                previous?.Dispose();
            }
            _surface = dark ? Color.FromArgb(45, 45, 45) : Color.White;
            _text = dark ? Color.FromArgb(241, 245, 249) : Color.FromArgb(15, 23, 42);
            _secondaryText = dark ? Color.FromArgb(203, 213, 225) : Color.FromArgb(51, 65, 85);
            _selectedSurface = dark ? Color.FromArgb(30, 58, 95) : Color.FromArgb(239, 246, 255);
            _border = dark ? Color.FromArgb(82, 82, 82) : Color.FromArgb(203, 213, 225);
            BackColor = _surface;
        }

        private void UpdateRoundedRegion()
        {
            if (Width <= 0 || Height <= 0) return;
            using (var path = RoundedPath(ClientRectangle, ScreenGeometry.Scale(9, _dpi)))
            {
                var previous = Region;
                Region = new Region(path);
                previous?.Dispose();
            }
        }

        private static bool SameItems(string[] left, string[] right)
        {
            if (ReferenceEquals(left, right)) return true;
            if (left == null || right == null || left.Length != right.Length) return false;
            for (int i = 0; i < left.Length; i++)
                if (!string.Equals(left[i], right[i], StringComparison.Ordinal)) return false;
            return true;
        }

        private static GraphicsPath RoundedPath(Rectangle rectangle, int radius)
        {
            var path = new GraphicsPath();
            int diameter = Math.Max(1, Math.Min(radius * 2,
                Math.Min(Math.Max(1, rectangle.Width), Math.Max(1, rectangle.Height))));
            path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
            path.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
            path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }

        protected override AccessibleObject CreateAccessibilityInstance() => new PopupAccessibility(this);

        private sealed class PopupAccessibility : ControlAccessibleObject
        {
            private readonly CompletionPopup _popup;
            internal PopupAccessibility(CompletionPopup popup) : base(popup) { _popup = popup; }
            public override AccessibleRole Role => AccessibleRole.List;
            public override int GetChildCount() => _popup._items.Length;
            public override AccessibleObject GetChild(int index) => index >= 0 && index < GetChildCount()
                ? new ItemAccessibility(_popup, index, _popup._items[index]) : null;
        }

        private sealed class ItemAccessibility : AccessibleObject
        {
            private readonly CompletionPopup _popup;
            private readonly int _index;
            private readonly string _word;
            internal ItemAccessibility(CompletionPopup popup, int index, string word)
            {
                _popup = popup;
                _index = index;
                _word = word;
            }
            public override string Name { get => _word; set { } }
            public override string DefaultAction => "Таклифни қабул қилиш";
            public override AccessibleRole Role => AccessibleRole.ListItem;
            public override AccessibleObject Parent => _popup.AccessibilityObject;
            public override Rectangle Bounds => _popup.RectangleToScreen(
                new Rectangle(1, _index * _popup.RowHeight, _popup.Width - 2, _popup.RowHeight));
            public override AccessibleStates State => AccessibleStates.Selectable |
                (_index == _popup.SelectedIndex ? AccessibleStates.Selected : AccessibleStates.None);
            public override void DoDefaultAction()
            {
                if (_popup.IsDisposed || !_popup.IsHandleCreated) return;
                _popup.BeginInvoke((Action)(() =>
                {
                    if (!_popup.IsDisposed && _popup.Visible && _index < _popup._items.Length &&
                        _popup._items[_index] == _word) _popup.Accepted?.Invoke(_word);
                }));
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                var font = _ownedFont;
                _ownedFont = null;
                Font = Control.DefaultFont;
                font?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
