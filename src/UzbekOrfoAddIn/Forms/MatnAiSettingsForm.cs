using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Threading.Tasks;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>Optional lifetime contract, independent of the prediction implementation.</summary>
    public interface IMatnAiSettingsWork
    {
        bool IsBusy { get; }
        event EventHandler BusyChanged;
        Task CancelAndWaitAsync();
    }

    public sealed class MatnAiSettingsForm : ModernForm
    {
        private readonly bool _initialLearning;
        private readonly Action<int, int, bool> _save;
        private readonly Action _clearLearning;
        private readonly Action _refreshIndex;
        private readonly Action<bool> _saveMetrics;
        private readonly NumericUpDown _minimum;
        private readonly NumericUpDown _count;
        private readonly ModernToggle _learning;
        private readonly CheckBox _metrics;
        private readonly Label _indexStatus;
        private readonly IMatnAiSettingsWork _backgroundWork;
        private ModernButton _saveButton;
        private bool _closing;
        private bool _closeReady;

        public bool MetricsConsent { get { return _metrics.Checked; } }
        protected override bool ReflowContent => false;

        public MatnAiSettingsForm(int minimum, int count, bool learningEnabled,
            Action<int, int, bool> save, Action clearLearning, Action refreshIndex,
            Control collectionsControl = null, bool metricsEnabled = false, Action<bool> saveMetrics = null)
        {
            _initialLearning = learningEnabled;
            _save = save ?? throw new ArgumentNullException(nameof(save));
            _clearLearning = clearLearning ?? throw new ArgumentNullException(nameof(clearLearning));
            _refreshIndex = refreshIndex ?? throw new ArgumentNullException(nameof(refreshIndex));
            _saveMetrics = saveMetrics;
            _backgroundWork = collectionsControl as IMatnAiSettingsWork;

            Title = "MatnAI — Созламалар";
            Size = new Size(900, 800);
            MinimumSize = new Size(600, 620);
            AllowResize = true;
            ShowMinimizeButton = false;
            _minimum = CreateNumber(Math.Max(2, Math.Min(15, minimum)), 2, 15, "Энг кам ҳарфлар сони");
            _count = CreateNumber(Math.Max(1, Math.Min(10, count)), 1, 10, "Таклифлар сони");
            _learning = new ModernToggle
            {
                IsOn = learningEnabled, ShowLabel = true, OnText = "Ёқилган", OffText = "Ўчирилган",
                Size = new Size(145, 32), AccessibleName = "Шахсий ўрганиш", BackColor = ThemeManager.Surface
            };
            _metrics = new CheckBox
            {
                Text = "Маҳаллий умумий статистикани ёқиш",
                AccessibleName = "Маҳаллий умумий статистикани ёқиш",
                Checked = metricsEnabled, AutoSize = true, ForeColor = ThemeManager.TextPrimary,
                Font = ThemeManager.FontBase, Margin = new Padding(0, 4, 0, 10)
            };
            _indexStatus = BodyLabel("Луғат сақланганда индекс автоматик янгиланади.");
            BuildContent(collectionsControl);
            BuildActions();
            if (_backgroundWork != null)
            {
                _backgroundWork.BusyChanged += BackgroundWorkChanged;
                BackgroundWorkChanged(this, EventArgs.Empty);
            }
        }

        private void BuildContent(Control collectionsControl)
        {
            ContentPanel.Padding = new Padding(20);
            var tabs = new SettingsTabs
            {
                Dock = DockStyle.Fill, Font = ThemeManager.FontBase,
                AccessibleName = "MatnAI созламалари"
            };
            var suggestions = NewPage("Таклифлар");
            var documents = NewPage("Ҳужжатлардан ўрганиш");
            var personal = NewPage("Шахсий ўрганиш");
            tabs.TabPages.AddRange(new[] { suggestions, documents, personal });
            var suggestionBody = Stack();
            Add(suggestionBody, PageHeading("Таклифларни ўзингизга мосланг"));
            Add(suggestionBody, BodyLabel("MatnAI офлайн ишлайди. Созламалар фақат шу компьютерда сақланади."));
            Add(suggestionBody, BuildSuggestionCard());
            Add(suggestionBody, BuildIndexCard());
            suggestions.Controls.Add(suggestionBody);
            if (collectionsControl != null)
            {
                // The collection control owns its scroll viewport. Nested automatic
                // scrolling can expand a docked child past the visible tab bounds.
                documents.AutoScroll = false;
                collectionsControl.Dock = DockStyle.Fill;
                documents.Controls.Add(collectionsControl);
            }
            else
            {
                var body = Stack();
                Add(body, BodyLabel("Ҳужжат тўпламлари хизмати ҳозир мавжуд эмас. Ҳужжатлардан ўрганиш учун MatnAI ойнасини қайта очинг."));
                documents.Controls.Add(body);
            }
            var personalBody = Stack();
            Add(personalBody, PageHeading("Ўрганиш ва махфийлик"));
            Add(personalBody, BodyLabel("Қайси маълумотлар сақланишини ўзингиз танланг. Барчаси шу компьютерда қолади."));
            Add(personalBody, BuildLearningCard());
            Add(personalBody, BodyLabel("Тозалаш дарҳол амалга ошади. «Бекор қилиш» бу амални қайтармайди. Ўрганишни ёқиш ёки ўчириш учун «Сақлаш»ни босинг."));
            var metricsBody = Stack();
            Add(metricsBody, _metrics);
            Add(metricsBody, BodyLabel("Ихтиёрий: фақат умумий сонлар ва ишлаш вақти шу компьютерда сақланади. Сўзлар, иборалар ва ҳужжат мазмуни статистикага киритилмайди; маълумот юборилмайди."));
            Add(personalBody, Card("Маҳаллий статистика", metricsBody));
            personal.Controls.Add(personalBody);
            ContentPanel.Controls.Add(tabs);
        }

        private static Label PageHeading(string text) => new Label
        {
            Text = text, AutoSize = true, Dock = DockStyle.Top,
            Font = ThemeManager.FontXLBold, ForeColor = ThemeManager.TextPrimary,
            Margin = new Padding(0, 4, 0, 12)
        };

        // Retain native tab keyboard navigation and accessibility, replacing only
        // the small classic tab headers. No custom focus or global shortcuts.
        private sealed class SettingsTabs : TabControl
        {
            private bool _fitting;
            internal SettingsTabs()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                DrawMode = TabDrawMode.OwnerDrawFixed;
                SizeMode = TabSizeMode.Fixed;
                Multiline = false;
                Padding = new Point(12, 8);
            }

            protected override void OnPaintBackground(PaintEventArgs e) => e.Graphics.Clear(ThemeManager.Background);

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.Clear(ThemeManager.Background);
                for (int index = 0; index < TabCount; index++)
                    OnDrawItem(new DrawItemEventArgs(e.Graphics, Font, GetTabRect(index), index,
                        index == SelectedIndex ? DrawItemState.Selected : DrawItemState.None));
            }

            protected override void OnResize(EventArgs e)
            {
                base.OnResize(e);
                FitHeaders();
            }

            protected override void OnFontChanged(EventArgs e)
            {
                base.OnFontChanged(e);
                FitHeaders();
            }

            protected override void OnSelectedIndexChanged(EventArgs e)
            {
                base.OnSelectedIndexChanged(e);
                FitHeaders();
            }

            private void FitHeaders()
            {
                if (_fitting) return;
                _fitting = true;
                try
                {
                    using (new DpiLayout.Context(IsHandleCreated ? DpiLayout.WindowContext(Handle) : IntPtr.Zero))
                    {
                        int height = Math.Max(44, Font.Height * 3);
                        var size = new Size(Math.Max(40, (ClientSize.Width - DpiLayout.Pixels(this, 40)) / Math.Max(3, TabCount)), height);
                        if (ItemSize != size) ItemSize = size;
                        Invalidate();
                    }
                }
                finally { _fitting = false; }
            }

            protected override void OnDrawItem(DrawItemEventArgs e)
            {
                if (e.Index < 0 || e.Index >= TabCount) return;
                using (var background = new SolidBrush(ThemeManager.Background)) e.Graphics.FillRectangle(background, e.Bounds);
                Rectangle bounds = Rectangle.Inflate(e.Bounds, -4, -4);
                bool selected = e.Index == SelectedIndex;
                int radius = Math.Max(6, Font.Height / 2), diameter = radius * 2;
                using (var path = new GraphicsPath())
                {
                    path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
                    path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
                    path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
                    path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
                    path.CloseFigure();
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    using (var fill = new SolidBrush(selected ? ThemeManager.Primary : ThemeManager.Surface))
                        e.Graphics.FillPath(fill, path);
                }
                TextRenderer.DrawText(e.Graphics, TabPages[e.Index].Text, Font, Rectangle.Inflate(bounds, -8, -3),
                    selected ? ThemeManager.TextOnPrimary : ThemeManager.TextPrimary,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
                if (selected && Focused && ShowFocusCues)
                    ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(bounds, -5, -5),
                        ThemeManager.TextOnPrimary, ThemeManager.Primary);
            }
        }

        private Control BuildSuggestionCard()
        {
            var body = Stack();
            Add(body, SettingRow("Энг кам ҳарфлар",
                "Таклиф кўрсатилиши учун ёзиладиган ҳарфлар сони.", _minimum));
            Add(body, SettingRow("Таклифлар сони",
                "Бир вақтда кўрсатиладиган энг яхши вариантлар сони.", _count));
            return Card("Таклифлар", body);
        }

        private Control BuildLearningCard()
        {
            var body = Stack();
            Add(body, SettingRow("Қабул қилинган таклифлардан ўрганиш",
                "Ихтиёрий: қабул қилинган сўз ва иборалар ҳамда уларни қабул қилиш сони шу компьютерда сақланади. Ҳужжатнинг тўлиқ матни сақланмайди ва юборилмайди.", _learning));
            var clear = Button("Ўрганишни тозалаш", ModernButton.ButtonStyle.Secondary);
            clear.AccessibleName = "Шахсий ўрганиш маълумотларини тозалаш";
            clear.Click += (s, e) =>
            {
                if (!ModernMessageBox.Confirm(
                    "Қабул қилинган сўз ва иборалар ҳамда уларнинг ҳисоби тозалансинми? Бу амални қайтариб бўлмайди. Луғатлар ва ҳужжат тўпламлари ўзгармайди.",
                    "MatnAI — Тасдиқ", "Тозалаш", "Бекор қилиш")) return;
                try
                {
                    _clearLearning();
                    ModernMessageBox.Success("Шахсий ўрганиш маълумотлари тозаланди.", "MatnAI");
                }
                catch (Exception ex) { ShowError(ex); }
            };
            Add(body, clear);
            return Card("Шахсий ўрганиш", body);
        }

        private Control BuildIndexCard()
        {
            var body = Stack();
            Add(body, _indexStatus);
            var refresh = Button("Ҳозир янгилаш", ModernButton.ButtonStyle.Secondary);
            refresh.AccessibleName = "Таклифлар индексини ҳозир янгилаш";
            refresh.Click += (s, e) =>
            {
                try
                {
                    _refreshIndex();
                    _indexStatus.Text = "Индекс янгиланмоқда — тайёр бўлгач таклифлар автоматик кўринади.";
                    _indexStatus.ForeColor = ThemeManager.Success;
                }
                catch (Exception ex) { ShowError(ex); }
            };
            Add(body, refresh);
            return Card("Луғат индекси", body);
        }

        private void BuildActions()
        {
            ActionBar.Height = 68;
            var cancel = Button("Бекор қилиш", ModernButton.ButtonStyle.Secondary);
            cancel.Size = new Size(155, 42);
            cancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            _saveButton = Button("Сақлаш", ModernButton.ButtonStyle.Primary);
            _saveButton.Size = new Size(155, 42);
            _saveButton.Click += SaveAndClose;
            ActionBar.Controls.Add(cancel);
            ActionBar.Controls.Add(_saveButton);
            ActionBar.Resize += (s, e) => LayoutActions(cancel);
            LayoutActions(cancel);
        }

        private void LayoutActions(Control cancel)
        {
            int y = (ActionBar.Height - _saveButton.Height) / 2;
            _saveButton.Location = new Point(ActionBar.Width - Px(24) - _saveButton.Width, y);
            cancel.Location = new Point(_saveButton.Left - Px(12) - cancel.Width, y);
        }

        private void SaveAndClose(object sender, EventArgs e)
        {
            if (_closing || (_backgroundWork != null && _backgroundWork.IsBusy)) return;
            if (!_initialLearning && _learning.IsOn && !ModernMessageBox.Confirm(
                "Қабул қилинган сўз ва иборалар ҳамда уларни қабул қилиш сони таклифларни яхшилаш учун фақат шу компьютерда сақлансинми?\n\nҲужжатнинг тўлиқ матни сақланмайди ва юборилмайди.",
                "MatnAI — Шахсий ўрганиш", "Розиман", "Бекор қилиш")) return;
            try
            {
                _save((int)_minimum.Value, (int)_count.Value, _learning.IsOn);
                _saveMetrics?.Invoke(MetricsConsent);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void BackgroundWorkChanged(object sender, EventArgs e)
        {
            if (!IsDisposed) _saveButton.Enabled = !_closing && !_backgroundWork.IsBusy;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_closeReady && (_closing || (_backgroundWork != null && _backgroundWork.IsBusy)))
            {
                e.Cancel = true;
                if (!_closing)
                {
                    _closing = true;
                    var result = DialogResult == DialogResult.None ? DialogResult.Cancel : DialogResult;
                    DialogResult = DialogResult.None;
                    ActionBar.Enabled = false;
                    ContentPanel.Enabled = false;
                    // Finish this closing event before attempting another Close.
                    BeginInvoke(new Action(() => CancelWorkAndCloseAsync(result)));
                }
            }
            base.OnFormClosing(e);
        }

        private async void CancelWorkAndCloseAsync(DialogResult result)
        {
            try
            {
                await _backgroundWork.CancelAndWaitAsync();
                if (IsDisposed) return;
                _closeReady = true;
                DialogResult = result;
                Close();
            }
            catch (Exception ex)
            {
                if (IsDisposed) return;
                _closing = false;
                ActionBar.Enabled = true;
                ContentPanel.Enabled = true;
                BackgroundWorkChanged(this, EventArgs.Empty);
                ShowError(ex);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _backgroundWork != null) _backgroundWork.BusyChanged -= BackgroundWorkChanged;
            base.Dispose(disposing);
        }

        private static TabPage NewPage(string text)
        {
            return new TabPage(text)
            {
                BackColor = ThemeManager.Background, ForeColor = ThemeManager.TextPrimary,
                Padding = new Padding(16), AutoScroll = true
            };
        }

        private static TableLayoutPanel Stack()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1, RowCount = 0, Margin = Padding.Empty, BackColor = Color.Transparent
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return panel;
        }

        private static void Add(TableLayoutPanel panel, Control control)
        {
            int row = panel.RowCount++;
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.Controls.Add(control, 0, row);
        }

        private static Control Card(string header, TableLayoutPanel body)
        {
            var card = new ModernCard
            {
                Header = header, Dock = DockStyle.Top, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0, 0, 0, 14),
                CornerRadius = ThemeManager.RadiusLG
            };
            card.Controls.Add(body);
            return card;
        }

        private static Control SettingRow(string title, string description, Control input)
        {
            var row = new TableLayoutPanel
            {
                AutoSize = true, Dock = DockStyle.Top, ColumnCount = 2, RowCount = 1,
                Margin = new Padding(0, 0, 0, 16)
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var labels = Stack();
            var heading = BodyLabel(title);
            heading.Font = ThemeManager.FontLGBold;
            heading.ForeColor = ThemeManager.TextPrimary;
            Add(labels, heading);
            Add(labels, BodyLabel(description));
            input.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            input.Margin = new Padding(12, 4, 0, 0);
            row.Controls.Add(labels, 0, 0);
            row.Controls.Add(input, 1, 0);
            return row;
        }

        private static NumericUpDown CreateNumber(int value, int minimum, int maximum, string name)
        {
            return new NumericUpDown
            {
                Minimum = minimum, Maximum = maximum, Value = value, Width = 104,
                Font = ThemeManager.FontLG, BackColor = ThemeManager.SurfaceElevated,
                ForeColor = ThemeManager.TextPrimary, BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center, AccessibleName = name
            };
        }

        private static ModernButton Button(string text, ModernButton.ButtonStyle style)
        {
            return new ModernButton
            {
                Text = text, AccessibleName = text, Style = style, Font = ThemeManager.FontBaseBold,
                Size = new Size(200, 38), MinimumSize = new Size(200, 38), Margin = new Padding(0, 4, 0, 10)
            };
        }

        private static Label BodyLabel(string text)
        {
            return new Label
            {
                Text = text, AutoSize = true, Dock = DockStyle.Top, Margin = new Padding(0, 0, 0, 10),
                Font = ThemeManager.FontBase, ForeColor = ThemeManager.TextSecondary,
                BackColor = Color.Transparent, UseMnemonic = false
            };
        }

        private static void ShowError(Exception exception)
        {
            ModernMessageBox.Error("Амал бажарилмади.\n\n" + exception.Message, "MatnAI");
        }
    }
}
