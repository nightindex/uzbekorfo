using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    public enum AddNewWordsMode
    {
        None = 0,
        File = 1,
        Folder = 2
    }

    /// <summary>
    /// Modern modal dialog to choose how to add new words (file or folder).
    /// </summary>
    public class AddNewWordsForm : ModernForm
    {
        public AddNewWordsMode SelectedMode { get; private set; } = AddNewWordsMode.None;

        public AddNewWordsForm()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96f, 96f);

            Title = "Сўз қўшиш";
            Size = new Size(960, 720);
            MinimumSize = new Size(860, 620);
            ShowMinimizeButton = false;
            AllowResize = false;

            BuildUI();
        }

        private void BuildUI()
        {
            ContentPanel.Padding = new Padding(ThemeManager.SpaceXL, ThemeManager.SpaceLG,
                                               ThemeManager.SpaceXL, ThemeManager.SpaceLG);

            var titleLabel = new Label
            {
                Text = "Сўз қўшиш",
                AutoSize = true,
                Font = new Font("Segoe UI", 20f, FontStyle.Bold),
                ForeColor = ThemeManager.TextPrimary
            };
            titleLabel.Margin = new Padding(0, 0, 0, 4);

            var subtitleLabel = new Label
            {
                Text = "Манбани танланг: файл ёки папка. Папка танланса, ички папкалар ҳам ҳисобга олинади.",
                AutoSize = true,
                Font = new Font("Segoe UI", 12.5f, FontStyle.Regular),
                ForeColor = ThemeManager.TextSecondary
            };
            subtitleLabel.Margin = new Padding(0, 0, 0, ThemeManager.SpaceLG);

            var cardsRow = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 400,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(0, ThemeManager.SpaceLG, 0, 0),
                BackColor = Color.Transparent
            };
            cardsRow.Margin = new Padding(0, 0, 0, ThemeManager.SpaceLG);
            cardsRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            cardsRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

            var fileCard = BuildOptionCard(
                title: "Файлдан қўшиш",
                description: "TXT, CSV, DOC/DOCX, XLS/XLSX ёки JSON файлидан сўзлар ва изоҳларни импорт қилади. Бир файлни тез қўшиш учун мос.",
                buttonText: "Файлдан қўшиш",
                mode: AddNewWordsMode.File);

            var folderCard = BuildOptionCard(
                title: "Папкадан қўшиш",
                description: "Танланган папка ва барча ички папкалардаги қўллаб-қувватланадиган файлларни йиғиб, бир марта импорт қилади. Катта коллекция учун қулай.",
                buttonText: "Папкадан қўшиш",
                mode: AddNewWordsMode.Folder);

            fileCard.Margin = new Padding(0, 0, ThemeManager.SpaceLG / 2, 0);
            folderCard.Margin = new Padding(ThemeManager.SpaceLG / 2, 0, 0, 0);

            cardsRow.Controls.Add(fileCard, 0, 0);
            cardsRow.Controls.Add(folderCard, 1, 0);

            var formatsLabel = new Label
            {
                Text = "Қўллаб-қувватланадиган форматлар: txt, dic, csv, doc, docx, xls, xlsx, json",
                AutoSize = true,
                Font = ThemeManager.FontBase,
                ForeColor = ThemeManager.TextSecondary
            };
            formatsLabel.Margin = new Padding(0, 0, 0, 0);

            var stack = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };
            stack.Controls.Add(titleLabel);
            stack.Controls.Add(subtitleLabel);
            stack.Controls.Add(cardsRow);
            stack.Controls.Add(formatsLabel);

            ContentPanel.Controls.Add(stack);

            // Action bar
            var btnCancel = new ModernButton
            {
                Text = "Бекор қилиш",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = ThemeManager.FontLGBold,
                Size = new Size(170, 46),
                Location = new Point(ActionBar.Width - 170 - ThemeManager.SpaceXL, 6),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            ActionBar.Controls.Add(btnCancel);
        }

        private ModernCard BuildOptionCard(string title, string description, string buttonText, AddNewWordsMode mode)
        {
            int pad = ThemeManager.SpaceLG;
            var card = new ModernCard
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                CornerRadius = ThemeManager.RadiusLG,
                MinimumSize = new Size(0, 380)
            };

            var iconBox = new PictureBox
            {
                Size = new Size(56, 56),
                Margin = new Padding(0, 0, 10, 0),
                BackColor = Color.Transparent
            };
            iconBox.Paint += (s, ev) => DrawCardIcon(ev.Graphics, iconBox.ClientRectangle, mode);

            var titleLabel = new Label
            {
                Text = title,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 14.5f, FontStyle.Bold),
                ForeColor = ThemeManager.TextPrimary,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Icon + Title side by side
            var headerRow = new TableLayoutPanel
            {
                AutoSize = false,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0, 0, 0, 16),
                Height = 56,
                BackColor = Color.Transparent
            };
            headerRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 66f));
            headerRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            headerRow.RowStyles.Add(new RowStyle(SizeType.Absolute, 56f));
            headerRow.Controls.Add(iconBox, 0, 0);
            headerRow.Controls.Add(titleLabel, 1, 0);

            var descLabel = new Label
            {
                Text = description,
                AutoSize = true,
                Font = new Font("Segoe UI", 12.5f, FontStyle.Regular),
                ForeColor = ThemeManager.TextSecondary,
                MaximumSize = new Size(420, 0),
                Margin = new Padding(0, 0, 0, 6)
            };

            var actionButton = new ModernButton
            {
                Text = buttonText,
                Style = ModernButton.ButtonStyle.Primary,
                Font = ThemeManager.FontLGBold,
                Size = new Size(200, 48),
                Anchor = AnchorStyles.Left | AnchorStyles.Bottom
            };
            actionButton.Click += (s, e) => SelectMode(mode);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(pad),
                BackColor = Color.Transparent
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // header (icon + title)
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // description
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // spacer
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // button

            layout.Controls.Add(headerRow, 0, 0);
            layout.Controls.Add(descLabel, 0, 1);
            layout.Controls.Add(actionButton, 0, 3);

            card.Controls.Add(layout);

            card.Resize += (s, e) =>
            {
                int maxWidth = Math.Max(240, card.ClientSize.Width - pad * 2);
                descLabel.MaximumSize = new Size(maxWidth, 0);
            };

            return card;
        }

        private void SelectMode(AddNewWordsMode mode)
        {
            SelectedMode = mode;
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>
        /// Draws a file or folder icon using GDI+ primitives.
        /// </summary>
        private static void DrawCardIcon(Graphics g, Rectangle rect, AddNewWordsMode mode)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            // Accent color for the icons
            var accent = Color.FromArgb(100, 149, 237); // Cornflower blue
            var accentDark = Color.FromArgb(70, 119, 207);

            if (mode == AddNewWordsMode.File)
            {
                // Draw a document/file icon — scaled to rect
                float x = rect.X + 6, y = rect.Y + 2;
                float w = rect.Width - 12, h = rect.Height - 4;
                float fold = w * 0.25f;

                using (var path = new GraphicsPath())
                {
                    path.AddLine(x, y, x + w - fold, y);
                    path.AddLine(x + w - fold, y, x + w, y + fold);
                    path.AddLine(x + w, y + fold, x + w, y + h);
                    path.AddLine(x + w, y + h, x, y + h);
                    path.CloseFigure();

                    using (var fill = new LinearGradientBrush(rect, accent, accentDark, 90f))
                        g.FillPath(fill, path);
                    using (var pen = new Pen(Color.FromArgb(60, 255, 255, 255), 1.4f))
                        g.DrawPath(pen, path);
                }

                // Corner fold
                using (var foldPath = new GraphicsPath())
                {
                    foldPath.AddLine(x + w - fold, y, x + w - fold, y + fold);
                    foldPath.AddLine(x + w - fold, y + fold, x + w, y + fold);
                    foldPath.CloseFigure();
                    using (var foldBrush = new SolidBrush(Color.FromArgb(80, 0, 0, 0)))
                        g.FillPath(foldBrush, foldPath);
                }

                // Lines on the document — proportional to height
                using (var linePen = new Pen(Color.FromArgb(160, 255, 255, 255), 1.8f))
                {
                    float lx = x + w * 0.14f;
                    float lw = w * 0.65f;
                    float lineY1 = y + h * 0.38f;
                    float lineGap = h * 0.12f;
                    g.DrawLine(linePen, lx, lineY1, lx + lw, lineY1);
                    g.DrawLine(linePen, lx, lineY1 + lineGap, lx + lw * 0.7f, lineY1 + lineGap);
                    g.DrawLine(linePen, lx, lineY1 + lineGap * 2, lx + lw * 0.85f, lineY1 + lineGap * 2);
                }
            }
            else
            {
                // Draw a folder icon — scaled to rect
                float x = rect.X + 3, y = rect.Y + 6;
                float w = rect.Width - 6, h = rect.Height - 10;
                float tabW = w * 0.4f, tabH = h * 0.16f;
                float r = 4f;

                // Folder tab
                using (var tabPath = new GraphicsPath())
                {
                    tabPath.AddArc(x, y, r * 2, r * 2, 180, 90);
                    tabPath.AddLine(x + r, y, x + tabW - 4, y);
                    tabPath.AddLine(x + tabW - 4, y, x + tabW + 4, y + tabH);
                    tabPath.AddLine(x + tabW + 4, y + tabH, x, y + tabH);
                    tabPath.CloseFigure();
                    using (var fill = new SolidBrush(accentDark))
                        g.FillPath(fill, tabPath);
                }

                // Folder body
                using (var bodyPath = new GraphicsPath())
                {
                    float by = y + tabH - 1;
                    float bh = h - tabH + 1;
                    bodyPath.AddArc(x, by, r * 2, r * 2, 180, 90);
                    bodyPath.AddArc(x + w - r * 2, by, r * 2, r * 2, 270, 90);
                    bodyPath.AddArc(x + w - r * 2, by + bh - r * 2, r * 2, r * 2, 0, 90);
                    bodyPath.AddArc(x, by + bh - r * 2, r * 2, r * 2, 90, 90);
                    bodyPath.CloseFigure();

                    using (var fill = new LinearGradientBrush(
                        new RectangleF(x, by, w, bh), accent, accentDark, 90f))
                        g.FillPath(fill, bodyPath);
                    using (var pen = new Pen(Color.FromArgb(60, 255, 255, 255), 1.4f))
                        g.DrawPath(pen, bodyPath);
                }
            }
        }

    }
}
