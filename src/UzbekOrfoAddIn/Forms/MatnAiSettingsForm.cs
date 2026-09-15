using System;
using System.Drawing;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>Theme-aware MatnAI preferences and local-data controls.</summary>
    public sealed class MatnAiSettingsForm : ModernForm
    {
        private readonly int _initialMinimum;
        private readonly int _initialCount;
        private readonly bool _initialLearning;
        private readonly Action<int, int, bool> _save;
        private readonly Action _clearLearning;
        private readonly Action _refreshIndex;
        private readonly NumericUpDown _minimum;
        private readonly NumericUpDown _count;
        private readonly ModernToggle _learning;
        private readonly Label _indexStatus;

        protected override bool ReflowContent => false;

        public MatnAiSettingsForm(int minimum, int count, bool learningEnabled,
            Action<int, int, bool> save, Action clearLearning, Action refreshIndex)
        {
            _initialMinimum = Math.Max(2, Math.Min(15, minimum));
            _initialCount = Math.Max(1, Math.Min(10, count));
            _initialLearning = learningEnabled;
            _save = save ?? throw new ArgumentNullException(nameof(save));
            _clearLearning = clearLearning ?? throw new ArgumentNullException(nameof(clearLearning));
            _refreshIndex = refreshIndex ?? throw new ArgumentNullException(nameof(refreshIndex));

            Title = "MatnAI — Созламалар";
            Size = new Size(720, 730);
            MinimumSize = new Size(560, 700);
            AllowResize = true;
            ShowMinimizeButton = false;

            _minimum = CreateNumber(Math.Max(3, _initialMinimum), 3, 15, "Энг кам ҳарфлар сони");
            _count = CreateNumber(_initialCount, 1, 10, "Таклифлар сони");
            _learning = new ModernToggle
            {
                IsOn = _initialLearning,
                ShowLabel = true,
                OnText = "Ёқилган",
                OffText = "Ўчирилган",
                Size = new Size(130, 30),
                AccessibleName = "Шахсий ўрганиш"
            };
            _indexStatus = BodyLabel("Луғат сақланганда индекс автоматик янгиланади.");

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
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 185f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 170f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 130f));

            var header = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            header.Controls.Add(new Label
            {
                Text = "Таклифларни ўзингизга мосланг",
                AutoSize = true,
                Font = ThemeManager.FontXLBold,
                ForeColor = ThemeManager.TextPrimary,
                Location = new Point(0, 2)
            });
            header.Controls.Add(new Label
            {
                Text = "MatnAI офлайн ишлайди. Бу созламалар фақат шу компьютерда сақланади.",
                AutoSize = false,
                Font = ThemeManager.FontBase,
                ForeColor = ThemeManager.TextSecondary,
                Location = new Point(0, 36),
                Size = new Size(640, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            });

            root.Controls.Add(header, 0, 0);
            root.Controls.Add(BuildSuggestionCard(), 0, 1);
            root.Controls.Add(BuildLearningCard(), 0, 2);
            root.Controls.Add(BuildIndexCard(), 0, 3);
            ContentPanel.Controls.Add(root);
        }

        private ModernCard BuildSuggestionCard()
        {
            var card = NewCard("ТАКЛИФЛАР");
            AddSettingRow(card, "Энг кам ҳарфлар",
                "Таклиф кўрсатилиши учун ёзиладиган ҳарфлар сони.", _minimum, 54);
            AddSettingRow(card, "Таклифлар сони",
                "Бир вақтда кўрсатиладиган энг яхши вариантлар сони.", _count, 112);
            return card;
        }

        private ModernCard BuildLearningCard()
        {
            var card = NewCard("ШАХСИЙ ЎРГАНИШ");
            var title = TitleLabel("Қабул қилинган сўзлардан ўрганиш");
            title.Location = new Point(20, 52);
            title.Size = new Size(390, 24);
            var description = BodyLabel("Фақат сўз ва қабул қилиш сони сақланади. Ҳужжат матни йиғилмайди ёки юборилмайди.");
            description.Location = new Point(20, 84);
            description.Size = new Size(420, 28);
            _learning.Location = new Point(500, 52);

            var clear = new ModernButton
            {
                Text = "Ўрганишни тозалаш",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = ThemeManager.FontBaseBold,
                Size = new Size(190, 34),
                Location = new Point(20, 116),
                AccessibleName = "Шахсий ўрганиш маълумотларини тозалаш"
            };
            clear.Click += (s, e) =>
            {
                if (!ModernMessageBox.Confirm(
                    "MatnAI қабул қилинган таклифлар ҳисобини тозаласинми? Луғатлар ўзгартирилмайди.",
                    "MatnAI — Тасдиқ", "Тозалаш", "Бекор қилиш")) return;
                _clearLearning();
                ModernMessageBox.Success("Шахсий ўрганиш маълумотлари тозаланди.", "MatnAI");
            };

            card.Controls.Add(title);
            card.Controls.Add(description);
            card.Controls.Add(_learning);
            card.Controls.Add(clear);
            card.Resize += (s, e) =>
            {
                _learning.Left = Math.Max(Px(260), card.ClientSize.Width - _learning.Width - Px(20));
                title.Width = Math.Max(Px(180), _learning.Left - Px(40));
                description.Width = Math.Max(Px(220), card.ClientSize.Width - Px(40));
            };
            return card;
        }

        private ModernCard BuildIndexCard()
        {
            var card = NewCard("ЛУҒАТ ИНДЕКСИ");
            _indexStatus.Location = new Point(20, 55);
            _indexStatus.Size = new Size(365, 42);

            var refresh = new ModernButton
            {
                Text = "Ҳозир янгилаш",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = ThemeManager.FontBaseBold,
                Size = new Size(180, 40),
                Location = new Point(450, 55),
                AccessibleName = "Таклифлар индексини ҳозир янгилаш"
            };
            refresh.Click += (s, e) =>
            {
                _refreshIndex();
                _indexStatus.Text = "Индекс янгиланмоқда — тайёр бўлгач таклифлар автоматик кўринади.";
                _indexStatus.ForeColor = ThemeManager.Success;
            };

            card.Controls.Add(_indexStatus);
            card.Controls.Add(refresh);
            card.Resize += (s, e) =>
            {
                refresh.Left = Math.Max(Px(260), card.ClientSize.Width - refresh.Width - Px(20));
                _indexStatus.Width = Math.Max(Px(210), refresh.Left - Px(40));
            };
            return card;
        }

        private void BuildActions()
        {
            ActionBar.Height = 68;
            var cancel = new ModernButton
            {
                Text = "Бекор қилиш",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = ThemeManager.FontLGBold,
                Size = new Size(150, 42)
            };
            cancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            var save = new ModernButton
            {
                Text = "Сақлаш",
                Style = ModernButton.ButtonStyle.Primary,
                Font = ThemeManager.FontLGBold,
                Size = new Size(150, 42)
            };
            save.Click += SaveAndClose;

            ActionBar.Controls.Add(cancel);
            ActionBar.Controls.Add(save);
            ActionBar.Resize += (s, e) =>
            {
                int y = (ActionBar.Height - save.Height) / 2;
                int right = ActionBar.Width - Px(ThemeManager.SpaceXL);
                save.Location = new Point(right - save.Width, y);
                cancel.Location = new Point(save.Left - Px(12) - cancel.Width, y);
            };
        }

        private void SaveAndClose(object sender, EventArgs e)
        {
            if (!_initialLearning && _learning.IsOn && !ModernMessageBox.Confirm(
                "Қабул қилинган сўзлар ва уларнинг сони таклифларни яхшилаш учун фақат шу компьютерда сақлансинми?\n\nҲужжат матни йиғилмайди ва юборилмайди.",
                "MatnAI — Шахсий ўрганиш", "Розиман", "Бекор қилиш")) return;

            _save((int)_minimum.Value, (int)_count.Value, _learning.IsOn);
            DialogResult = DialogResult.OK;
            Close();
        }

        private static NumericUpDown CreateNumber(int value, int minimum, int maximum, string accessibleName)
        {
            return new NumericUpDown
            {
                Minimum = minimum,
                Maximum = maximum,
                Value = value,
                Width = 82,
                Height = 32,
                Font = ThemeManager.FontLG,
                BackColor = ThemeManager.SurfaceElevated,
                ForeColor = ThemeManager.TextPrimary,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center,
                AccessibleName = accessibleName
            };
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

        private static void AddSettingRow(Control card, string title, string description,
            Control input, int y)
        {
            var titleLabel = TitleLabel(title);
            titleLabel.Location = new Point(20, y);
            titleLabel.Size = new Size(390, 22);
            var descriptionLabel = BodyLabel(description);
            descriptionLabel.Location = new Point(20, y + 24);
            descriptionLabel.Size = new Size(430, 28);
            input.Location = new Point(548, y + 5);
            card.Controls.Add(titleLabel);
            card.Controls.Add(descriptionLabel);
            card.Controls.Add(input);
            card.Resize += (s, e) =>
            {
                int margin = DpiLayout.Pixels(card, 20);
                int gap = DpiLayout.Pixels(card, 40);
                input.Left = Math.Max(DpiLayout.Pixels(card, 250), card.ClientSize.Width - input.Width - margin);
                titleLabel.Width = Math.Max(DpiLayout.Pixels(card, 180), input.Left - gap);
                descriptionLabel.Width = Math.Max(DpiLayout.Pixels(card, 220), input.Left - gap);
            };
        }

        private static Label TitleLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = false,
                Font = ThemeManager.FontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                BackColor = Color.Transparent
            };
        }

        private static Label BodyLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = false,
                Font = ThemeManager.FontBase,
                ForeColor = ThemeManager.TextSecondary,
                BackColor = Color.Transparent
            };
        }
    }
}
