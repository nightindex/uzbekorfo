using System;
using System.Drawing;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>Modern quick-start guide for MatnAI word completion.</summary>
    public sealed class MatnAiHelpForm : ModernForm
    {
        protected override bool ReflowContent => false;

        public MatnAiHelpForm()
        {
            Title = "MatnAI — Қисқа қўлланма";
            Size = new Size(760, 740);
            MinimumSize = new Size(560, 725);
            AllowResize = true;
            ShowMinimizeButton = false;

            BuildContent();
            BuildActions();
        }

        private void BuildContent()
        {
            ContentPanel.Padding = new Padding(ThemeManager.SpaceXL, ThemeManager.SpaceLG,
                ThemeManager.SpaceXL, ThemeManager.SpaceLG);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                BackColor = Color.Transparent,
                Margin = Padding.Empty
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 82f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 150f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 215f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 145f));

            var header = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            header.Controls.Add(new Label
            {
                Text = "Тез ёзишни бошланг",
                AutoSize = true,
                Font = ThemeManager.FontXLBold,
                ForeColor = ThemeManager.TextPrimary,
                Location = new Point(0, 2)
            });
            header.Controls.Add(new Label
            {
                Text = "MatnAI сўзни тугатишга ёрдам беради ва матнни фақат сиз қабул қилганда ўзгартиради.",
                AutoSize = false,
                Font = ThemeManager.FontBase,
                ForeColor = ThemeManager.TextSecondary,
                Location = new Point(0, 36),
                Size = new Size(680, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            });

            root.Controls.Add(header, 0, 0);
            root.Controls.Add(BuildStartCard(), 0, 1);
            root.Controls.Add(BuildShortcutCard(), 0, 2);
            root.Controls.Add(BuildPrivacyCard(), 0, 3);
            ContentPanel.Controls.Add(root);
        }

        private static ModernCard BuildStartCard()
        {
            var card = NewCard("БОШЛАШ");
            AddStep(card, "1", "Таклифларни ёқиш",
                "Тасмадаги «Таклифларни ёқиш» тугмасини фаоллаштиринг.", 46);
            AddStep(card, "2", "Сўз ёзиш",
                "Камида иккита ҳарф ёзинг. Энг яхши давоми курсор ёнида кулранг матн бўлиб кўринади.", 93);
            return card;
        }

        private static ModernCard BuildShortcutCard()
        {
            var card = NewCard("ТЕЗКОР ТУГМАЛАР");
            AddShortcut(card, "Tab", "Кулранг таклифни қабул қилиш", 47);
            AddShortcut(card, "↓; сўнг ↑ / ↓", "Таклифлар рўйхатини очиш ва танлаш", 85);
            AddShortcut(card, "Esc", "Таклифни ёпиш", 123);
            AddShortcut(card, "Ctrl + Alt + →", "Таклифни қабул қилишнинг муқобил тугмаси", 161);
            return card;
        }

        private static ModernCard BuildPrivacyCard()
        {
            var card = NewCard("МАХФИЙЛИК ВА ХАВФСИЗЛИК");
            var badge = new Label
            {
                Text = "✓",
                AutoSize = false,
                Size = new Size(36, 36),
                Location = new Point(20, 58),
                Font = ThemeManager.FontXLBold,
                ForeColor = ThemeManager.Success,
                BackColor = ThemeManager.SuccessLight,
                TextAlign = ContentAlignment.MiddleCenter
            };
            var text = new Label
            {
                Text = "MatnAI офлайн ишлайди. Ҳужжат матни йиғилмайди ва юборилмайди. Шахсий ўрганиш ихтиёрий; ҳимояланган матн, жадвал ва кузатилаётган ўзгаришларда таклифлар кўрсатилмайди.",
                AutoSize = false,
                Location = new Point(70, 53),
                Size = new Size(580, 70),
                Font = ThemeManager.FontBase,
                ForeColor = ThemeManager.TextSecondary,
                BackColor = Color.Transparent
            };
            card.Controls.Add(badge);
            card.Controls.Add(text);
            card.Resize += (s, e) => text.Width = Math.Max(DpiLayout.Pixels(card, 260),
                card.ClientSize.Width - DpiLayout.Pixels(card, 90));
            return card;
        }

        private void BuildActions()
        {
            ActionBar.Height = 68;
            var close = new ModernButton
            {
                Text = "Тушунарли",
                Style = ModernButton.ButtonStyle.Primary,
                Font = ThemeManager.FontLGBold,
                Size = new Size(160, 42)
            };
            close.Click += (s, e) => { DialogResult = DialogResult.OK; Close(); };
            ActionBar.Controls.Add(close);
            ActionBar.Resize += (s, e) =>
            {
                int y = (ActionBar.Height - close.Height) / 2;
                close.Location = new Point(ActionBar.Width - Px(ThemeManager.SpaceXL) - close.Width, y);
            };
        }

        private static void AddStep(Control card, string number, string title, string body, int y)
        {
            var badge = new Label
            {
                Text = number,
                AutoSize = false,
                Size = new Size(30, 30),
                Location = new Point(20, y),
                Font = ThemeManager.FontBaseBold,
                ForeColor = ThemeManager.Primary,
                BackColor = ThemeManager.PrimaryLight,
                TextAlign = ContentAlignment.MiddleCenter
            };
            var titleLabel = new Label
            {
                Text = title,
                AutoSize = false,
                Location = new Point(64, y - 2),
                Size = new Size(570, 22),
                Font = ThemeManager.FontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                BackColor = Color.Transparent
            };
            var bodyLabel = new Label
            {
                Text = body,
                AutoSize = false,
                Location = new Point(64, y + 22),
                Size = new Size(570, 22),
                Font = ThemeManager.FontBase,
                ForeColor = ThemeManager.TextSecondary,
                BackColor = Color.Transparent
            };
            card.Controls.Add(badge);
            card.Controls.Add(titleLabel);
            card.Controls.Add(bodyLabel);
            card.Resize += (s, e) =>
            {
                titleLabel.Width = Math.Max(DpiLayout.Pixels(card, 240),
                    card.ClientSize.Width - DpiLayout.Pixels(card, 84));
                bodyLabel.Width = Math.Max(DpiLayout.Pixels(card, 240),
                    card.ClientSize.Width - DpiLayout.Pixels(card, 84));
            };
        }

        private static void AddShortcut(Control card, string shortcut, string action, int y)
        {
            var key = new Label
            {
                Text = shortcut,
                AutoSize = false,
                Size = new Size(180, 32),
                Location = new Point(20, y),
                Font = ThemeManager.FontMono,
                ForeColor = ThemeManager.Primary,
                BackColor = ThemeManager.SurfaceHover,
                TextAlign = ContentAlignment.MiddleCenter,
                AccessibleName = shortcut
            };
            var description = new Label
            {
                Text = action,
                AutoSize = false,
                Location = new Point(220, y),
                Size = new Size(410, 32),
                Font = ThemeManager.FontLG,
                ForeColor = ThemeManager.TextPrimary,
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft
            };
            card.Controls.Add(key);
            card.Controls.Add(description);
            card.Resize += (s, e) => description.Width = Math.Max(DpiLayout.Pixels(card, 220),
                card.ClientSize.Width - DpiLayout.Pixels(card, 240));
        }

        private static ModernCard NewCard(string header)
        {
            return new ModernCard
            {
                Header = header,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, ThemeManager.SpaceMD),
                CornerRadius = ThemeManager.RadiusLG
            };
        }
    }
}
