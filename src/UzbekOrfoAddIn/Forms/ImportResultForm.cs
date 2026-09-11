using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>
    /// Modal dialog that shows the result summary after a folder/file import.
    /// Displays word count, definition count, processed files, and any failures.
    /// </summary>
    public class ImportResultForm : ModernForm
    {
        /// <summary>The result type determines the accent bar color and badge icon.</summary>
        public enum ResultType { Success, Warning, Info }

        private ResultType _type;
        private int _addedWords;
        private int _addedDefinitions;
        private int _processedFiles;
        private int _totalDictionaryWords;
        private int _skippedInMain;

        private ImportResultForm(ResultType type, int addedWords, int addedDefinitions,
                                  int processedFiles, int totalDictionaryWords, int skippedInMain = 0)
        {
            _type = type;
            _addedWords = addedWords;
            _addedDefinitions = addedDefinitions;
            _processedFiles = processedFiles;
            _totalDictionaryWords = totalDictionaryWords;
            _skippedInMain = skippedInMain;

            Title = "Импорт натижаси";
            Size = new Size(640, 430);
            MinimumSize = new Size(560, 390);
            ShowMinimizeButton = false;
            AllowResize = false;

            BuildUI();
        }

        private void BuildUI()
        {
            // --- Action bar with OK button ---
            ActionBar.Height = 64;

            var btnOk = new ModernButton
            {
                Text = "Тушундим",
                Style = ModernButton.ButtonStyle.Primary,
                Font = ThemeManager.FontLGBold,
                Size = new Size(160, 44)
            };
            btnOk.Click += (s, e) =>
            {
                DialogResult = DialogResult.OK;
            };
            ActionBar.Controls.Add(btnOk);

            ActionBar.Resize += (s, e) =>
            {
                int y = (ActionBar.Height - btnOk.Height) / 2;
                int right = ActionBar.Width - Px(ThemeManager.SpaceXL);
                btnOk.Location = new Point(right - btnOk.Width, y);
            };

            // --- Content area (custom paint + stat cards) ---
            ContentPanel.Padding = new Padding(ThemeManager.SpaceXL, ThemeManager.SpaceLG,
                                               ThemeManager.SpaceXL, ThemeManager.SpaceLG);

            // Badge + summary text via custom paint
            ContentPanel.AutoScroll = false;
            ContentPanel.Paint += ContentPanel_Paint;

            // Stat cards in a 2x2 table, positioned below the header
            int pad = ThemeManager.SpaceXL;
            int gap = ThemeManager.SpaceSM;
            int cardH = 68;

            var table = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 2,
                Location = new Point(pad, 100),
                AutoSize = false,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, cardH + gap));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, cardH + gap));

            table.Size = new Size(
                ContentPanel.ClientSize.Width - pad * 2,
                2 * (cardH + gap));

            // Words added
            table.Controls.Add(CreateStatCard(
                _addedWords.ToString("N0"),
                "та янги сўз",
                ThemeManager.Primary), 0, 0);

            // Definitions added
            table.Controls.Add(CreateStatCard(
                _addedDefinitions.ToString("N0"),
                "та изоҳ",
                ThemeManager.Accent), 1, 0);

            // Processed files
            table.Controls.Add(CreateStatCard(
                _processedFiles.ToString("N0"),
                "файл ишланди",
                ThemeManager.Success), 0, 1);

            // Total dictionary words
            table.Controls.Add(CreateStatCard(
                FormatLargeNumber(_totalDictionaryWords),
                "жами сўзлар",
                ThemeManager.Primary), 1, 1);

            ContentPanel.Controls.Add(table);
        }

        private void ContentPanel_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Badge circle
            int badgeSize = Px(48);
            int badgeX = Px(ThemeManager.SpaceXL);
            int badgeY = Px(ThemeManager.SpaceLG);

            Color badgeColor = GetAccentColor();
            using (var badgeBrush = new SolidBrush(badgeColor))
            {
                g.FillEllipse(badgeBrush, badgeX, badgeY, badgeSize, badgeSize);
            }

            // Badge icon (checkmark, warning or info)
            string iconChar = _type == ResultType.Success ? "✓"
                            : _type == ResultType.Warning ? "!" : "i";
            using (var iconFont = new Font("Segoe UI", 18f, FontStyle.Bold))
            using (var iconBrush = new SolidBrush(Color.White))
            {
                var iconSize = g.MeasureString(iconChar, UiFont(iconFont));
                float ix = badgeX + (badgeSize - iconSize.Width) / 2f;
                float iy = badgeY + (badgeSize - iconSize.Height) / 2f;
                g.DrawString(iconChar, UiFont(iconFont), iconBrush, ix, iy);
            }

            // Title text next to badge
            string title = GetTitleText();
            int textX = badgeX + badgeSize + Px(ThemeManager.SpaceMD);
            int textMaxW = ContentPanel.ClientSize.Width - textX - Px(ThemeManager.SpaceXL);

            using (var titleFont = new Font("Segoe UI", 16f, FontStyle.Bold))
            using (var titleBrush = new SolidBrush(ThemeManager.TextPrimary))
            {
                var titleSize = g.MeasureString(title, UiFont(titleFont));
                int titleY = badgeY + (badgeSize - (int)titleSize.Height) / 2;
                g.DrawString(title, UiFont(titleFont), titleBrush,
                    new RectangleF(textX, titleY, textMaxW, Px(36)));
            }

            // Subtitle below badge row
            string subtitle = GetSubtitleText();
            if (!string.IsNullOrEmpty(subtitle))
            {
                int subY = badgeY + badgeSize + Px(6);
                using (var subFont = new Font("Segoe UI", 11f, FontStyle.Regular))
                using (var subBrush = new SolidBrush(ThemeManager.TextSecondary))
                {
                    g.DrawString(subtitle, UiFont(subFont), subBrush,
                        new RectangleF(badgeX, subY, textMaxW + badgeSize + Px(ThemeManager.SpaceMD), Px(24)),
                        new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap });
                }
            }
        }

        private string GetTitleText()
        {
            if (_addedWords == 0 && _addedDefinitions == 0)
            {
                if (_skippedInMain > 0)
                    return "Сўзлар асосий луғатда мавжуд";
                return "Янги сўз топилмади";
            }
            return "Импорт муваффақиятли тугади!";
        }

        private string GetSubtitleText()
        {
            if (_addedWords == 0 && _addedDefinitions == 0)
            {
                if (_skippedInMain > 0)
                    return $"{_skippedInMain:N0} та сўз асосий луғатда аллақачон мавжуд — қайтадан қўшиш шарт эмас.";
                return "Танланган файллардан янги сўз ёки изоҳ топилмади.";
            }
            return $"{_processedFiles} файлдан жами {_addedWords + _addedDefinitions} та маълумот юкланди.";
        }

        private Color GetAccentColor()
        {
            switch (_type)
            {
                case ResultType.Success: return ThemeManager.Success;
                case ResultType.Warning: return ThemeManager.Warning;
                default: return ThemeManager.Primary;
            }
        }

        /// <summary>Paint the accent strip at the very top of the form.</summary>
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using (var brush = new SolidBrush(GetAccentColor()))
            {
                e.Graphics.FillRectangle(brush, 0, 0, Width, Px(4));
            }
        }

        /// <summary>Creates a mini stat card with a large number and label.</summary>
        private Control CreateStatCard(string number, string label,
                                               Color accentColor)
        {
            var card = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, ThemeManager.SpaceSM, ThemeManager.SpaceSM),
                BackColor = ThemeManager.Surface
            };
            card.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                int r = Px(ThemeManager.RadiusSM);
                var rect = new Rectangle(0, 0, card.Width - 1, card.Height - 1);
                using (var path = RoundedRect(rect, r))
                {
                    // Card background
                    using (var bg = new SolidBrush(ThemeManager.Surface))
                        g.FillPath(bg, path);

                    // Subtle border
                    using (var pen = new Pen(ThemeManager.Border))
                        g.DrawPath(pen, path);
                }

                // Accent bar on left
                using (var accentBrush = new SolidBrush(accentColor))
                    g.FillRectangle(accentBrush, 0, Px(6), Px(4), card.Height - Px(12));

                // Number — auto-scale font for large numbers
                float numFontSize = number.Length > 9 ? 13f
                                  : number.Length > 7 ? 15f
                                  : number.Length > 5 ? 17f
                                  : 20f;
                using (var numFont = new Font("Segoe UI", numFontSize, FontStyle.Bold))
                using (var numBrush = new SolidBrush(accentColor))
                {
                    var numSize = g.MeasureString(number, UiFont(numFont));
                    float numY = Px(6) + (Px(30) - numSize.Height) / 2f;
                    g.DrawString(number, UiFont(numFont), numBrush, Px(16), Math.Max(Px(4), numY));
                }

                // Label
                using (var lblFont = new Font("Segoe UI", 10.5f, FontStyle.Regular))
                using (var lblBrush = new SolidBrush(ThemeManager.TextSecondary))
                {
                    g.DrawString(label, UiFont(lblFont), lblBrush, Px(16), Px(42));
                }
            };

            return card;
        }

        private static GraphicsPath RoundedRect(Rectangle rect, int radius)
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

        // ===================== Static API =====================

        /// <summary>
        /// Shows the import result dialog. Determines the result type automatically.
        /// </summary>
        /// <summary>
        /// Formats a large number with abbreviated suffix for display.
        /// e.g. 1,234,567 → "1,234,567" (with comma grouping).
        /// </summary>
        private static string FormatLargeNumber(int value)
        {
            return value.ToString("N0");
        }

        public static void ShowResult(int addedWords, int addedDefinitions,
                                       int processedFiles, int totalDictionaryWords,
                                       int skippedInMain = 0)
        {
            ResultType type;
            if (addedWords == 0 && addedDefinitions == 0)
                type = ResultType.Info;
            else
                type = ResultType.Success;

            using (var form = new ImportResultForm(type, addedWords, addedDefinitions,
                                                   processedFiles, totalDictionaryWords, skippedInMain))
            {
                form.ShowDialog();
            }
        }
    }
}
