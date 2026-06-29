using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.UI.Controls
{
    /// <summary>
    /// A modern text box with rounded borders, placeholder text, focus highlight,
    /// and theme-aware styling. Wraps a standard TextBox with custom painting.
    /// Uses a Label overlay for dark-theme-safe placeholder rendering.
    /// </summary>
    public class ModernTextBox : UserControl
    {
        private readonly TextBox _innerTextBox;
        private readonly Label _placeholderLabel;
        private string _placeholder = "";
        private bool _focused;
        private int _cornerRadius = 8;
        private Image _prefixIcon;

        /// <summary>
        /// The actual text content.
        /// </summary>
        public override string Text
        {
            get => _innerTextBox.Text;
            set => _innerTextBox.Text = value;
        }

        /// <summary>
        /// Placeholder text shown when the box is empty.
        /// </summary>
        public string Placeholder
        {
            get => _placeholder;
            set
            {
                _placeholder = value;
                if (_placeholderLabel != null)
                    _placeholderLabel.Text = value;
                UpdatePlaceholderVisibility();
                Invalidate();
            }
        }

        /// <summary>
        /// Optional icon shown at the left side (e.g., search icon).
        /// </summary>
        public Image PrefixIcon
        {
            get => _prefixIcon;
            set { _prefixIcon = value; UpdateLayout(); Invalidate(); }
        }

        /// <summary>
        /// Corner radius for the border.
        /// </summary>
        public int CornerRadius
        {
            get => _cornerRadius;
            set { _cornerRadius = value; Invalidate(); }
        }

        /// <summary>
        /// Whether the inner text box accepts multi-line input.
        /// </summary>
        public bool Multiline
        {
            get => _innerTextBox.Multiline;
            set => _innerTextBox.Multiline = value;
        }

        /// <summary>
        /// Whether the text is read-only.
        /// </summary>
        public bool ReadOnly
        {
            get => _innerTextBox.ReadOnly;
            set => _innerTextBox.ReadOnly = value;
        }

        /// <summary>
        /// Fires when the text changes.
        /// </summary>
        public new event EventHandler TextChanged;

        public ModernTextBox()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw, true);

            Size = new Size(240, 36);
            Padding = new Padding(12, 6, 12, 6);

            _innerTextBox = new TextBox
            {
                BorderStyle = BorderStyle.None,
                BackColor = ThemeManager.Background,
                ForeColor = ThemeManager.TextPrimary,
                Font = ThemeManager.FontBase
            };

            // --- Placeholder overlay label ---
            // Sits on top of _innerTextBox in the z-order; hides when user types or focuses.
            _placeholderLabel = new Label
            {
                AutoSize = false,
                BackColor = ThemeManager.Background,   // match textbox bg exactly
                ForeColor = ThemeManager.TextDisabled,
                Font = ThemeManager.FontBase,
                Cursor = Cursors.IBeam,
                TextAlign = ContentAlignment.MiddleLeft,
                UseMnemonic = false
            };
            _placeholderLabel.Click += (s, e) => _innerTextBox.Focus();
            _placeholderLabel.MouseDown += (s, e) => _innerTextBox.Focus();

            _innerTextBox.GotFocus += (s, e) => { _focused = true; UpdatePlaceholderVisibility(); Invalidate(); };
            _innerTextBox.LostFocus += (s, e) => { _focused = false; UpdatePlaceholderVisibility(); Invalidate(); };
            _innerTextBox.TextChanged += (s, e) => { TextChanged?.Invoke(this, e); UpdatePlaceholderVisibility(); Invalidate(); };

            // Add both — order matters: _placeholderLabel added AFTER _innerTextBox so it's
            // on top in the z-order (BringToFront ensures it).
            Controls.Add(_innerTextBox);
            Controls.Add(_placeholderLabel);
            _placeholderLabel.BringToFront();

            UpdateLayout();
        }

        /// <summary>
        /// Show/hide the placeholder label based on focus and text content.
        /// The placeholder is visible only when the TextBox is empty AND not focused.
        /// </summary>
        private void UpdatePlaceholderVisibility()
        {
            if (_placeholderLabel == null) return;
            bool show = string.IsNullOrEmpty(_innerTextBox.Text)
                        && !_focused
                        && !string.IsNullOrEmpty(_placeholder);
            _placeholderLabel.Visible = show;
        }

        private void UpdateLayout()
        {
            if (_innerTextBox == null) return;
            int leftPad = _prefixIcon != null ? 34 : 12;
            int textW = Width - leftPad - 12;

            _innerTextBox.Location = new Point(leftPad, (Height - _innerTextBox.Height) / 2);
            _innerTextBox.Width = textW;

            if (_placeholderLabel != null)
            {
                _placeholderLabel.Location = _innerTextBox.Location;
                _placeholderLabel.Size = new Size(textW, _innerTextBox.Height);
                _placeholderLabel.Text = _placeholder;
                _placeholderLabel.Font = _innerTextBox.Font;
                _placeholderLabel.BackColor = _innerTextBox.BackColor;
                UpdatePlaceholderVisibility();
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UpdateLayout();
        }

        /// <summary>
        /// When the user sets Font on the ModernTextBox, propagate to inner controls.
        /// </summary>
        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            if (_innerTextBox != null) _innerTextBox.Font = Font;
            if (_placeholderLabel != null) _placeholderLabel.Font = Font;
            UpdateLayout();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);

            // Background fill
            using (var path = CreateRoundedRect(rect, _cornerRadius))
            using (var brush = new SolidBrush(ThemeManager.Background))
            {
                g.FillPath(brush, path);
            }

            // Border
            var borderColor = _focused ? ThemeManager.BorderFocus : ThemeManager.Border;
            int borderWidth = _focused ? 2 : 1;
            using (var path = CreateRoundedRect(rect, _cornerRadius))
            using (var pen = new Pen(borderColor, borderWidth))
            {
                g.DrawPath(pen, path);
            }

            // Prefix icon
            if (_prefixIcon != null)
            {
                g.DrawImage(_prefixIcon, new Rectangle(10, (Height - 16) / 2, 16, 16));
            }

            // NOTE: Placeholder is now rendered via _placeholderLabel overlay,
            // NOT painted here, because child controls paint on top of OnPaint.
        }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            _innerTextBox.Focus();
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
