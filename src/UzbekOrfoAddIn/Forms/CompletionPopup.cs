using System;
using System.Drawing;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>Painted non-activating popup: typing focus remains in Word.</summary>
    public sealed class CompletionPopup : Form
    {
        private string[] _items = new string[0];
        public int SelectedIndex { get; private set; }
        private int RowHeight => Math.Max(30, Font.Height + 12);
        public event Action<string> Accepted;
        public event Action Dismissed;
        public string Selected => _items.Length == 0 ? null : _items[SelectedIndex];
        protected override bool ShowWithoutActivation => true;
        protected override CreateParams CreateParams
        {
            get { var cp = base.CreateParams; cp.ExStyle |= 0x08000000 | 0x80; return cp; }
        }
        public CompletionPopup()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 10);
            DoubleBuffered = true;
            AccessibleName = "MatnAI сўз таклифлари";
        }
        public void Present(string[] items, Rectangle caret, IWin32Window owner)
        {
            bool changed = !object.ReferenceEquals(items, _items);
            _items = items;
            if (changed) SelectedIndex = 0;
            Width = 320;
            Height = (items.Length + 1) * RowHeight + 2;
            var area = Screen.FromRectangle(caret).WorkingArea;
            Location = new Point(Math.Max(area.Left, Math.Min(caret.Left, area.Right - Width)),
                caret.Bottom + Height <= area.Bottom ? caret.Bottom : Math.Max(area.Top, caret.Top - Height));
            AccessibleDescription = string.Join(", ", items);
            if (!Visible) Show(owner);
            if (changed)
            {
                AccessibilityNotifyClients(AccessibleEvents.Reorder, -1);
                Invalidate();
            }
        }
        public void MoveSelection(int direction)
        {
            if (_items.Length == 0 || !Visible) return;
            SelectedIndex = (SelectedIndex + direction + _items.Length) % _items.Length;
            AccessibilityNotifyClients(AccessibleEvents.Selection, SelectedIndex);
            Invalidate();
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x21) { m.Result = new IntPtr(3); return; } // MA_NOACTIVATE
            base.WndProc(ref m);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.DrawRectangle(Pens.SlateGray, 0, 0, Width - 1, Height - 1);
            for (int i = 0; i < _items.Length; i++)
            {
                if (i == SelectedIndex) e.Graphics.FillRectangle(SystemBrushes.Info, 1, i * RowHeight + 1, Width - 2, RowHeight);
                TextRenderer.DrawText(e.Graphics, (i + 1) + ". " + _items[i], Font,
                    new Rectangle(8, i * RowHeight + 4, Width - 16, RowHeight), Color.Black,
                    TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
            TextRenderer.DrawText(e.Graphics, "Ctrl+Alt+Right: қабул қилиш     Esc: ёпиш", Font,
                new Rectangle(8, _items.Length * RowHeight + 4, Width - 16, RowHeight), Color.DimGray,
                TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;
            int index = e.Y / RowHeight;
            if (index < _items.Length) Accepted?.Invoke(_items[index]);
            else Dismissed?.Invoke();
        }

        protected override AccessibleObject CreateAccessibilityInstance() => new PopupAccessibility(this);
        private sealed class PopupAccessibility : ControlAccessibleObject
        {
            private readonly CompletionPopup _popup;
            public PopupAccessibility(CompletionPopup popup) : base(popup) { _popup = popup; }
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
            public ItemAccessibility(CompletionPopup popup, int index, string word) { _popup = popup; _index = index; _word = word; }
            public override string Name { get => _word; set { } }
            public override string DefaultAction => "Таклифни қабул қилиш";
            public override AccessibleRole Role => AccessibleRole.ListItem;
            public override AccessibleObject Parent => _popup.AccessibilityObject;
            public override Rectangle Bounds => _popup.RectangleToScreen(new Rectangle(1, _index * _popup.RowHeight, _popup.Width - 2, _popup.RowHeight));
            public override AccessibleStates State => AccessibleStates.Selectable |
                (_index == _popup.SelectedIndex ? AccessibleStates.Selected : AccessibleStates.None);
            public override void DoDefaultAction()
            {
                if (_popup.IsDisposed || !_popup.IsHandleCreated) return;
                _popup.BeginInvoke((Action)(() =>
                {
                    if (!_popup.IsDisposed && _popup.Visible && _index < _popup._items.Length && _popup._items[_index] == _word)
                        _popup.Accepted?.Invoke(_word);
                }));
            }
        }
    }
}
