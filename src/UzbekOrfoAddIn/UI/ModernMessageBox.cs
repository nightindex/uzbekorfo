using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.UI
{
    /// <summary>
    /// A polished, modern themed message box replacing MessageBox.Show().
    /// Features: colored accent strip, large icon in a circle badge, Uzbek button labels.
    /// 
    /// Usage:
    ///   ModernMessageBox.Success("Хато топилмади!");
    ///   ModernMessageBox.Warning("3 та хато топилди.");
    ///   bool yes = ModernMessageBox.Confirm("Ўчирилсинми?");
    /// </summary>
    public class ModernMessageBox : ModernForm
    {
        public enum MessageType { Info, Success, Warning, Error }

        private int ACCENT_STRIP_HEIGHT => Px(4);
        private const int ICON_CIRCLE_SIZE = 56;

        private DialogResult _result = DialogResult.Cancel;
        private MessageType _messageType;

        private ModernMessageBox(string message, string title, MessageType type,
                                  bool showCancel, string okText, string cancelText,
                                  Image titleIcon = null)
        {
            _messageType = type;
            Title = title;
            AllowResize = false;
            ShowMinimizeButton = false;
            Size = new Size(540, 290);
            MinimumSize = new Size(420, 230);
            MaximizeBox = false;

            if (titleIcon != null)
                TitleIcon = titleIcon;

            BuildContent(message, type, showCancel, okText, cancelText);
        }

        // =============================================================
        //  ACCENT STRIP — colored bar at the very top of the window
        // =============================================================

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            // Draw accent strip below the title bar
            var accentColor = GetIconColor(_messageType);
            using (var brush = new SolidBrush(accentColor))
            {
                // The title bar is painted by ModernForm at y=0..TITLE_BAR_HEIGHT
                // We paint a thin accent strip right below it
                e.Graphics.FillRectangle(brush, 0, 0, Width, ACCENT_STRIP_HEIGHT);
            }
        }

        // =============================================================
        //  CONTENT LAYOUT
        // =============================================================

        private void BuildContent(string message, MessageType type,
                                   bool showCancel, string okText, string cancelText)
        {
            int pad = ThemeManager.SpaceXL;
            int iconAreaWidth = ICON_CIRCLE_SIZE + 16; // circle + gap to text

            // === Draw icon circle directly on ContentPanel (no z-order issues) ===
            ContentPanel.Padding = new Padding(pad, ThemeManager.SpaceLG, pad, ThemeManager.SpaceSM);
            ContentPanel.Paint += (s, pe) =>
            {
                var g = pe.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                int cx = Px(pad);
                int cy = (ContentPanel.Height - Px(ICON_CIRCLE_SIZE)) / 2;

                // Circle background (light tint of the icon color)
                var iconColor = GetIconColor(type);
                var circleColor = Color.FromArgb(35, iconColor);
                using (var brush = new SolidBrush(circleColor))
                {
                    g.FillEllipse(brush, cx, cy, Px(ICON_CIRCLE_SIZE), Px(ICON_CIRCLE_SIZE));
                }

                // Icon symbol centered in circle
                var iconText = GetIcon(type);
                using (var iconFont = new Font("Segoe UI", 22f, FontStyle.Regular))
                using (var brush = new SolidBrush(iconColor))
                {
                    var sf = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center
                    };
                    g.DrawString(iconText, UiFont(iconFont), brush,
                        new RectangleF(cx, cy, Px(ICON_CIRCLE_SIZE), Px(ICON_CIRCLE_SIZE)), sf);
                }
            };

            // === Message label — positioned to the right of the painted icon ===
            var msgLabel = new Label
            {
                Text = message,
                Font = ThemeManager.FontXL,
                ForeColor = ThemeManager.TextPrimary,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.Transparent
            };

            ContentPanel.Controls.Add(msgLabel);

            // Position label to the right of icon, vertically centered on the icon
            ContentPanel.Layout += (s, e) =>
            {
                int leftEdge = Px(pad + iconAreaWidth);
                int rightEdge = ContentPanel.ClientSize.Width - Px(pad);
                int textWidth = Math.Max(1, rightEdge - leftEdge);
                int textHeight = Math.Max(Px(ICON_CIRCLE_SIZE), msgLabel.GetPreferredSize(new Size(textWidth, 0)).Height);
                msgLabel.SetBounds(leftEdge, Px(pad), textWidth, textHeight);
                int neededHeight = textHeight + Px(pad) * 2;
                bool overflow = neededHeight > ContentPanel.ClientSize.Height;
                ContentPanel.AutoScroll = overflow;
                ContentPanel.AutoScrollMinSize = overflow ? new Size(0, neededHeight) : Size.Empty;
            };

            // === Action bar buttons ===
            ActionBar.Height = 68;
            ActionBar.Padding = new Padding(pad, 0, pad, 0);

            // Measure button widths to fit text
            int btnHeight = 42;
            int btnPadX = 48; // horizontal padding inside button
            int okWidth, cancelWidth = 0;
            okWidth = TextRenderer.MeasureText(okText, ThemeManager.FontLG).Width + btnPadX;
            okWidth = Math.Max(okWidth, 150);
            if (showCancel)
            {
                cancelWidth = TextRenderer.MeasureText(cancelText, ThemeManager.FontLG).Width + btnPadX;
                cancelWidth = Math.Max(cancelWidth, 150);
            }

            var btnOk = new ModernButton
            {
                Text = okText,
                Style = ModernButton.ButtonStyle.Primary,
                Font = ThemeManager.FontLG,
                Size = new Size(okWidth, btnHeight)
            };
            btnOk.Click += (s, e) =>
            {
                _result = DialogResult.OK;
                Close();
            };

            if (showCancel)
            {
                var btnCancel = new ModernButton
                {
                    Text = cancelText,
                    Style = ModernButton.ButtonStyle.Secondary,
                    Font = ThemeManager.FontLG,
                    Size = new Size(cancelWidth, btnHeight)
                };
                btnCancel.Click += (s, e) =>
                {
                    _result = DialogResult.Cancel;
                    Close();
                };

                ActionBar.Resize += (s, e) =>
                {
                    int y = (ActionBar.Height - btnOk.Height) / 2;
                    int right = ActionBar.Width - Px(pad);
                    btnOk.Location = new Point(right - btnOk.Width, y);
                    btnCancel.Location = new Point(btnOk.Left - Px(12) - btnCancel.Width, y);
                };
                ActionBar.Controls.Add(btnOk);
                ActionBar.Controls.Add(btnCancel);
            }
            else
            {
                ActionBar.Resize += (s, e) =>
                {
                    int y = (ActionBar.Height - btnOk.Height) / 2;
                    int right = ActionBar.Width - Px(pad);
                    btnOk.Location = new Point(right - btnOk.Width, y);
                };
                ActionBar.Controls.Add(btnOk);
            }

            // Ensure dialog is wide enough for the buttons + icon area
            int totalBtnWidth = okWidth + (showCancel ? cancelWidth + 12 : 0) + pad * 2 + 20;
            int minWidth = Math.Max(540, totalBtnWidth);
            if (Size.Width < minWidth)
                Size = new Size(minWidth, Size.Height);

            // Auto-size height based on message length
            using (var g = CreateGraphics())
            {
                int textAreaW = Size.Width - ICON_CIRCLE_SIZE - 100;
                var measured = g.MeasureString(message, ThemeManager.FontXL, textAreaW);
                int neededHeight = Math.Max(290, (int)measured.Height + 210);
                Size = new Size(Size.Width, Math.Min(neededHeight, 450));
            }
        }

        // =============================================================
        //  ICON HELPERS
        // =============================================================

        private static string GetIcon(MessageType type)
        {
            switch (type)
            {
                case MessageType.Success: return "\u2714"; // ✔
                case MessageType.Warning: return "\u26A0"; // ⚠
                case MessageType.Error:   return "\u2716"; // ✖
                default:                  return "\u2139"; // ℹ
            }
        }

        private static Color GetIconColor(MessageType type)
        {
            switch (type)
            {
                case MessageType.Success: return Color.FromArgb(34, 197, 94);   // green
                case MessageType.Warning: return Color.FromArgb(245, 158, 11);  // amber
                case MessageType.Error:   return Color.FromArgb(239, 68, 68);   // red
                default:                  return ThemeManager.Primary;           // blue
            }
        }

        // =============================================================
        //  PUBLIC STATIC API
        // =============================================================

        /// <summary>Shows a modern message dialog with a single OK button.</summary>
        public static void Show(string message, string title = "\u040e\u0437\u0431\u0435\u043a \u041e\u0440\u0444\u043e",
                                 MessageType type = MessageType.Info,
                                 string okText = "\u0422\u0443\u0448\u0443\u043d\u0434\u0438\u043c",
                                 Image icon = null)
        {
            using (var dlg = new ModernMessageBox(message, title, type, false, okText, "", icon))
            {
                dlg.ShowDialog();
            }
        }

        /// <summary>Shows a success message dialog.</summary>
        public static void Success(string message, string title = "\u040e\u0437\u0431\u0435\u043a \u041e\u0440\u0444\u043e",
                                    Image icon = null)
        {
            Show(message, title, MessageType.Success, "\u042f\u0445\u0448\u0438", icon);
        }

        /// <summary>Shows a warning message dialog.</summary>
        public static void Warning(string message, string title = "\u040e\u0437\u0431\u0435\u043a \u041e\u0440\u0444\u043e \u2014 \u041e\u0433\u043e\u04b3\u043b\u0430\u043d\u0442\u0438\u0440\u0438\u0448",
                                    Image icon = null)
        {
            Show(message, title, MessageType.Warning, "\u0422\u0443\u0448\u0443\u043d\u0434\u0438\u043c", icon);
        }

        /// <summary>Shows an error message dialog.</summary>
        public static void Error(string message, string title = "\u040e\u0437\u0431\u0435\u043a \u041e\u0440\u0444\u043e \u2014 \u0425\u0430\u0442\u043e\u043b\u0438\u043a",
                                  Image icon = null)
        {
            Show(message, title, MessageType.Error, "\u0422\u0443\u0448\u0443\u043d\u0434\u0438\u043c", icon);
        }

        /// <summary>Shows a Yes/No confirmation dialog. Returns true if user clicks Yes.</summary>
        public static bool Confirm(string message, string title = "\u040e\u0437\u0431\u0435\u043a \u041e\u0440\u0444\u043e \u2014 \u0422\u0430\u0441\u0434\u0438\u049b",
                                    string yesText = "\u04b2\u0430", string noText = "\u0419\u045e\u049b",
                                    Image icon = null)
        {
            using (var dlg = new ModernMessageBox(message, title, MessageType.Warning, true, yesText, noText, icon))
            {
                dlg.ShowDialog();
                return dlg._result == DialogResult.OK;
            }
        }

        /// <summary>
        /// Shows a 3-button dialog (Yes / No / Cancel).
        /// Returns DialogResult.Yes, DialogResult.No, or DialogResult.Cancel.
        /// </summary>
        public static DialogResult ConfirmOrCancel(
            string message,
            string title = "\u040e\u0437\u0431\u0435\u043a \u041e\u0440\u0444\u043e \u2014 \u0422\u0430\u0441\u0434\u0438\u049b",
            string yesText = "\u04b2\u0430",
            string noText = "\u0419\u045e\u049b",
            string cancelText = "\u0411\u0435\u043a\u043e\u0440",
            Image icon = null)
        {
            using (var dlg = new ModernMessageBox(message, title, MessageType.Info, false, "", "", icon))
            {
                // Remove default OK button and build 3-button layout manually
                dlg.ActionBar.Controls.Clear();

                int pad = ThemeManager.SpaceXL;
                int btnHeight = 42;
                int btnPadX = 48;

                int yesW = Math.Max(TextRenderer.MeasureText(yesText, ThemeManager.FontLG).Width + btnPadX, 130);
                int noW = Math.Max(TextRenderer.MeasureText(noText, ThemeManager.FontLG).Width + btnPadX, 130);
                int cancelW = Math.Max(TextRenderer.MeasureText(cancelText, ThemeManager.FontLG).Width + btnPadX, 130);

                var btnYes = new ModernButton
                {
                    Text = yesText,
                    Style = ModernButton.ButtonStyle.Primary,
                    Font = ThemeManager.FontLG,
                    Size = new Size(yesW, btnHeight)
                };
                btnYes.Click += (s, e) => { dlg._result = DialogResult.Yes; dlg.Close(); };

                var btnNo = new ModernButton
                {
                    Text = noText,
                    Style = ModernButton.ButtonStyle.Secondary,
                    Font = ThemeManager.FontLG,
                    Size = new Size(noW, btnHeight)
                };
                btnNo.Click += (s, e) => { dlg._result = DialogResult.No; dlg.Close(); };

                var btnCancel = new ModernButton
                {
                    Text = cancelText,
                    Style = ModernButton.ButtonStyle.Secondary,
                    Font = ThemeManager.FontLG,
                    Size = new Size(cancelW, btnHeight)
                };
                btnCancel.Click += (s, e) => { dlg._result = DialogResult.Cancel; dlg.Close(); };

                dlg.ActionBar.Controls.Add(btnYes);
                dlg.ActionBar.Controls.Add(btnNo);
                dlg.ActionBar.Controls.Add(btnCancel);

                dlg.ActionBar.Resize += (s, e) =>
                {
                    int y = (dlg.ActionBar.Height - btnCancel.Height) / 2;
                    int right = dlg.ActionBar.Width - dlg.Px(pad);
                    btnCancel.Location = new Point(right - btnCancel.Width, y);
                    btnNo.Location = new Point(btnCancel.Left - dlg.Px(12) - btnNo.Width, y);
                    btnYes.Location = new Point(btnNo.Left - dlg.Px(12) - btnYes.Width, y);
                };

                // Ensure dialog is wide enough
                int totalBtnWidth = yesW + noW + cancelW + 24 + pad * 2 + 20;
                int minWidth = Math.Max(540, totalBtnWidth);
                if (dlg.Size.Width < minWidth)
                    dlg.Size = new Size(minWidth, dlg.Size.Height);

                dlg._result = DialogResult.Cancel; // default if closed via X
                dlg.ShowDialog();
                return dlg._result;
            }
        }
    }
}
