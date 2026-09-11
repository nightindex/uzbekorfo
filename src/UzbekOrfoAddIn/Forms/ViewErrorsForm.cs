using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>
    /// Modern dialog for viewing and navigating spelling and grammar errors.
    /// Displays errors in two tabs: Spelling and Grammar.
    /// </summary>
    public class ViewErrorsForm : ModernForm
    {
        private readonly List<ErrorEntry> _spellingErrors;
        private readonly List<ErrorEntry> _grammarErrors;
        private readonly Action<ErrorEntry, string> _onReplace;
        private readonly Action<ErrorEntry> _onIgnore;
        private readonly Func<ErrorEntry, AddWordResult> _onAddToDict;

        // Controls
        private Panel _pillBar;
        private Label _pillSpelling;
        private Label _pillGrammar;
        private Panel _gridContainer;
        private DataGridView _spellingGrid;
        private DataGridView _grammarGrid;
        private Label _detailLabel;
        private Label _ruleLabel;
        private Label _statusLabel;
        private ModernTextBox _searchBox;
        private ModernButton _btnReplace;
        private ModernButton _btnIgnore;
        private ModernButton _btnAddDict;
        private ModernButton _btnGoTo;
        private ModernButton _btnPrev;
        private ModernButton _btnNext;

        private int _currentIndex = -1;
        private int _activeTab = 0; // 0 = Spelling, 1 = Grammar
        private List<ErrorEntry> _filteredSpelling = new List<ErrorEntry>();
        private List<ErrorEntry> _filteredGrammar = new List<ErrorEntry>();
        private bool _isInitialized;
        private bool _populatingGrids;

        // 498px of left actions + 170px replacement panel + 32px padding.
        protected override int GetMinimumActionBarWidth() => 720;

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            BeginInvoke((Action)(() =>
            {
                if (IsDisposed || !Visible) return;
                if (ActiveGrid.SelectedRows.Count > 0)
                    _currentIndex = ActiveGrid.SelectedRows[0].Index;
                UpdateDetail();
            }));
        }

        // Pill fonts
        private static readonly Font _fPill = new Font("Segoe UI Semibold", 11f, FontStyle.Bold);
        private static readonly Font _fPillR = new Font("Segoe UI", 11f, FontStyle.Regular);

        // +1pt fonts for error text readability
        private Font _fontLG1;
        private Font _fontLGBold1;
        private Font _fontBase1;

        /// <summary>
        /// Creates a new ViewErrors dialog with separate spelling and grammar tabs.
        /// </summary>
        public ViewErrorsForm(
            List<ErrorEntry> spellingErrors,
            List<ErrorEntry> grammarErrors,
            Action<ErrorEntry, string> onReplace,
            Action<ErrorEntry> onIgnore,
            Func<ErrorEntry, AddWordResult> onAddToDict)
        {
            _spellingErrors = spellingErrors ?? new List<ErrorEntry>();
            _grammarErrors = grammarErrors ?? new List<ErrorEntry>();
            _filteredSpelling = _spellingErrors.Where(e => !e.IsResolved).ToList();
            _filteredGrammar = _grammarErrors.Where(e => !e.IsResolved).ToList();
            _onReplace = onReplace;
            _onIgnore = onIgnore;
            _onAddToDict = onAddToDict;

            Title = "Хатоларни кўриш";
            Size = new Size(1080, 750);
            MinimumSize = new Size(750, 500);
            Font = ThemeManager.FontLG;

            _fontLG1 = new Font(ThemeManager.FontFamily, 12.25f, FontStyle.Regular);
            _fontLGBold1 = new Font(ThemeManager.FontFamily, 12.25f, FontStyle.Bold);
            _fontBase1 = new Font(ThemeManager.FontFamily, 10.75f, FontStyle.Regular);

            BuildUI();
            PopulateGrids();
            _isInitialized = true;
        }

        /// <summary>
        /// Whether the grammar tab is currently active.
        /// </summary>
        private bool IsGrammarTab => _activeTab == 1;

        /// <summary>
        /// The currently active grid.
        /// </summary>
        private DataGridView ActiveGrid => IsGrammarTab ? _grammarGrid : _spellingGrid;

        /// <summary>
        /// The currently active filtered error list.
        /// </summary>
        private List<ErrorEntry> ActiveErrors => IsGrammarTab ? _filteredGrammar : _filteredSpelling;

        private void BuildUI()
        {
            // === Top area: search + count ===
            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 96,
                Padding = new Padding(0, 0, 0, 0)
            };

            _searchBox = new ModernTextBox
            {
                Placeholder = "Қидириш...",
                Font = _fontLG1,
                Location = new Point(0, 4),
                Size = new Size(280, 38)
            };
            _searchBox.TextChanged += (s, e) => FilterErrors();

            _statusLabel = new Label
            {
                Text = "",
                AutoSize = true,
                ForeColor = ThemeManager.TextPrimary,
                Font = _fontLGBold1,
                Location = new Point(300, 10)
            };

            // === Pill tab bar ===
            _pillBar = new Panel
            {
                Location = new Point(0, 50),
                Size = new Size(600, 42),
                BackColor = Color.Transparent
            };

            _pillSpelling = CreatePill($"Имло хатолари ({_filteredSpelling.Count})", 0);
            _pillGrammar = CreatePill($"Грамматик хатолар ({_filteredGrammar.Count})", 1);

            // Measure and position pills
            int tw1 = TextRenderer.MeasureText(_pillSpelling.Text, UiFont(_fPill)).Width + Px(32);
            int tw2 = TextRenderer.MeasureText(_pillGrammar.Text, UiFont(_fPill)).Width + Px(32);
            _pillSpelling.Size = new Size(tw1, Px(38));
            _pillGrammar.Size = new Size(tw2, Px(38));
            _pillSpelling.Location = new Point(0, 2);
            _pillGrammar.Location = new Point(tw1 + Px(10), Px(2));

            _pillBar.Controls.Add(_pillSpelling);
            _pillBar.Controls.Add(_pillGrammar);

            topPanel.Controls.Add(_searchBox);
            topPanel.Controls.Add(_statusLabel);
            topPanel.Controls.Add(_pillBar);

            // === Grid container (swap grids based on active pill) ===
            _gridContainer = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeManager.Background
            };

            _spellingGrid = CreateGrid();
            _spellingGrid.Columns.Add("colNum", "№");
            _spellingGrid.Columns.Add("colWord", "Хато сўз");
            _spellingGrid.Columns.Add("colContext", "Контекст");
            _spellingGrid.Columns.Add("colSuggestion", "Таклиф");
            _spellingGrid.Columns.Add("colPara", "§");
            ConfigureGridColumns(_spellingGrid);

            _grammarGrid = CreateGrid();
            _grammarGrid.Columns.Add("colNum", "№");
            _grammarGrid.Columns.Add("colWord", "Сўз");
            _grammarGrid.Columns.Add("colMessage", "Қоида");
            _grammarGrid.Columns.Add("colSuggestion", "Таклиф");
            _grammarGrid.Columns.Add("colPara", "§");
            ConfigureGridColumns(_grammarGrid);

            _spellingGrid.Dock = DockStyle.Fill;
            _grammarGrid.Dock = DockStyle.Fill;
            _grammarGrid.Visible = false;

            _gridContainer.Controls.Add(_spellingGrid);
            _gridContainer.Controls.Add(_grammarGrid);

            // === Detail panel (bottom) ===
            var detailPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 68,
                BackColor = ThemeManager.Surface,
                Padding = new Padding(ThemeManager.SpaceLG, ThemeManager.SpaceSM,
                                      ThemeManager.SpaceLG, ThemeManager.SpaceSM)
            };

            _detailLabel = new Label
            {
                Text = "Сўзни танланг...",
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 26,
                ForeColor = ThemeManager.TextPrimary,
                Font = _fontLGBold1
            };

            _ruleLabel = new Label
            {
                Text = "",
                AutoSize = false,
                Dock = DockStyle.Fill,
                ForeColor = ThemeManager.TextSecondary,
                Font = _fontBase1
            };

            detailPanel.Controls.Add(_ruleLabel);
            detailPanel.Controls.Add(_detailLabel);

            var detailSep = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 1,
                BackColor = ThemeManager.Border
            };

            // === Action bar buttons ===
            ActionBar.Height = 72;
            ActionBar.Padding = new Padding(ThemeManager.SpaceLG, 0, ThemeManager.SpaceLG, 0);

            var leftFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = false,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            var rightPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 170,
                Padding = new Padding(0)
            };

            int btnH = 40;
            var btnFont = ThemeManager.FontLG;
            var btnMargin = new Padding(0, 0, 8, 0);

            _btnPrev = new ModernButton
            {
                Text = "◀",
                Font = btnFont,
                Style = ModernButton.ButtonStyle.Ghost,
                Size = new Size(42, btnH),
                Margin = new Padding(0, 0, 2, 0)
            };
            _btnPrev.Click += (s, e) => NavigateError(-1);

            _btnNext = new ModernButton
            {
                Text = "▶",
                Font = btnFont,
                Style = ModernButton.ButtonStyle.Ghost,
                Size = new Size(42, btnH),
                Margin = btnMargin
            };
            _btnNext.Click += (s, e) => NavigateError(1);

            _btnIgnore = new ModernButton
            {
                Text = "Ўтказиш",
                Font = btnFont,
                Style = ModernButton.ButtonStyle.Secondary,
                Size = new Size(130, btnH),
                Margin = btnMargin
            };
            _btnIgnore.Click += BtnIgnore_Click;

            _btnAddDict = new ModernButton
            {
                Text = "Луғатга",
                Font = btnFont,
                Style = ModernButton.ButtonStyle.Secondary,
                Size = new Size(120, btnH),
                Margin = btnMargin
            };
            _btnAddDict.Click += BtnAddDict_Click;

            _btnGoTo = new ModernButton
            {
                Text = "Кўрсатиш",
                Font = btnFont,
                Style = ModernButton.ButtonStyle.Secondary,
                Size = new Size(130, btnH),
                Margin = btnMargin
            };
            _btnGoTo.Click += BtnGoTo_Click;

            _btnReplace = new ModernButton
            {
                Text = "Алмаштириш",
                Font = btnFont,
                Style = ModernButton.ButtonStyle.Primary,
                Size = new Size(160, btnH),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            _btnReplace.Click += BtnReplace_Click;

            leftFlow.Resize += (s, e) =>
            {
                int cy = (leftFlow.Height - Px(btnH)) / 2;
                foreach (Control c in leftFlow.Controls)
                    c.Margin = new Padding(c.Margin.Left, cy, c.Margin.Right, 0);
            };
            rightPanel.Resize += (s, e) =>
            {
                _btnReplace.Location = new Point(
                    rightPanel.Width - _btnReplace.Width,
                    (rightPanel.Height - _btnReplace.Height) / 2);
            };

            leftFlow.Controls.Add(_btnPrev);
            leftFlow.Controls.Add(_btnNext);
            leftFlow.Controls.Add(_btnIgnore);
            leftFlow.Controls.Add(_btnAddDict);
            leftFlow.Controls.Add(_btnGoTo);

            rightPanel.Controls.Add(_btnReplace);

            ActionBar.Controls.Add(leftFlow);
            ActionBar.Controls.Add(rightPanel);

            // === Assemble layout ===
            ContentPanel.Padding = new Padding(ThemeManager.SpaceLG, ThemeManager.SpaceSM,
                                               ThemeManager.SpaceLG, 0);
            ContentPanel.Controls.Add(_gridContainer);
            ContentPanel.Controls.Add(detailSep);
            ContentPanel.Controls.Add(detailPanel);
            ContentPanel.Controls.Add(topPanel);
        }

        private DataGridView CreateGrid()
        {
            var grid = new DpiDataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None,
                BackgroundColor = ThemeManager.Background,
                GridColor = ThemeManager.Border,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                EnableHeadersVisualStyles = false,
                RowTemplate = { Height = 40 },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = ThemeManager.Background,
                    ForeColor = ThemeManager.TextPrimary,
                    SelectionBackColor = ThemeManager.SelectedRow,
                    SelectionForeColor = ThemeManager.TextPrimary,
                    Font = _fontLG1,
                    Padding = new Padding(6)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = ThemeManager.Surface,
                    ForeColor = ThemeManager.TextSecondary,
                    Font = _fontLGBold1,
                    Padding = new Padding(6),
                    Alignment = DataGridViewContentAlignment.MiddleLeft
                },
                ColumnHeadersHeight = 40
            };

            grid.SelectionChanged += Grid_SelectionChanged;
            grid.CellDoubleClick += Grid_CellDoubleClick;
            return grid;
        }

        private void ConfigureGridColumns(DataGridView grid)
        {
            grid.Columns[0].Width = 40;
            grid.Columns[0].FillWeight = 8;
            grid.Columns[1].FillWeight = 20;
            grid.Columns[2].FillWeight = 42;
            grid.Columns[3].FillWeight = 22;
            grid.Columns[4].FillWeight = 8;
        }

        // =====================================================================
        //  DATA
        // =====================================================================

        private void PopulateGrids()
        {
            _populatingGrids = true;
            try
            {
                PopulateSpellingGrid();
                PopulateGrammarGrid();
                UpdateStatusAndButtons();
                if (_spellingGrid.Rows.Count > 0)
                    _spellingGrid.Rows[0].Selected = true;
            }
            finally { _populatingGrids = false; }
            if (_isInitialized && Visible && ActiveGrid.SelectedRows.Count > 0)
            {
                _currentIndex = ActiveGrid.SelectedRows[0].Index;
                UpdateDetail();
            }
        }
        private void PopulateSpellingGrid()
        {
            _spellingGrid.Rows.Clear();
            for (int i = 0; i < _filteredSpelling.Count; i++)
            {
                var err = _filteredSpelling[i];
                if (err == null) continue;

                string best = err.BestSuggestion;
                _spellingGrid.Rows.Add(
                    (i + 1).ToString(),
                    err.Word ?? "",
                    TruncateContext(err.Context, 50),
                    best ?? "-",
                    err.ParagraphIndex.ToString()
                );
            }
        }

        private void PopulateGrammarGrid()
        {
            _grammarGrid.Rows.Clear();
            for (int i = 0; i < _filteredGrammar.Count; i++)
            {
                var err = _filteredGrammar[i];
                if (err == null) continue;

                string best = err.BestSuggestion;
                _grammarGrid.Rows.Add(
                    (i + 1).ToString(),
                    err.Word ?? "",
                    TruncateContext(err.Message ?? err.Context, 55),
                    best ?? "-",
                    err.ParagraphIndex.ToString()
                );
            }
        }

        private void UpdateStatusAndButtons()
        {
            int total = _filteredSpelling.Count + _filteredGrammar.Count;
            _statusLabel.Text = $"Жами: {total} та хато (имло: {_filteredSpelling.Count}, грамматика: {_filteredGrammar.Count})";

            // Update pill labels with counts
            _pillSpelling.Text = $"Имло хатолари ({_filteredSpelling.Count})";
            _pillGrammar.Text = $"Грамматик хатолар ({_filteredGrammar.Count})";

            // Resize pills to fit new text
            int tw1 = TextRenderer.MeasureText(_pillSpelling.Text, UiFont(_fPill)).Width + Px(32);
            int tw2 = TextRenderer.MeasureText(_pillGrammar.Text, UiFont(_fPill)).Width + Px(32);
            _pillSpelling.Size = new Size(tw1, Px(38));
            _pillGrammar.Size = new Size(tw2, Px(38));
            _pillGrammar.Location = new Point(tw1 + Px(10), Px(2));

            _pillSpelling.Invalidate();
            _pillGrammar.Invalidate();

            // Hide "Add to Dictionary" for grammar tab (not applicable)
            _btnAddDict.Visible = !IsGrammarTab;
        }

        private void FilterErrors()
        {
            if (!_isInitialized) return;

            string query = _searchBox?.Text?.Trim().ToLowerInvariant() ?? "";

            if (string.IsNullOrEmpty(query))
            {
                _filteredSpelling = _spellingErrors.Where(e => e != null && !e.IsResolved).ToList();
                _filteredGrammar = _grammarErrors.Where(e => e != null && !e.IsResolved).ToList();
            }
            else
            {
                _filteredSpelling = _spellingErrors
                    .Where(e => e != null && !e.IsResolved && MatchesQuery(e, query)).ToList();
                _filteredGrammar = _grammarErrors
                    .Where(e => e != null && !e.IsResolved && MatchesQuery(e, query)).ToList();
            }

            PopulateGrids();
        }

        private static bool MatchesQuery(ErrorEntry e, string query)
        {
            return (e.Word?.ToLowerInvariant().Contains(query) ?? false) ||
                   (e.Context?.ToLowerInvariant().Contains(query) ?? false) ||
                   (e.Message?.ToLowerInvariant().Contains(query) ?? false);
        }

        // =====================================================================
        //  EVENTS
        // =====================================================================

        private void Grid_SelectionChanged(object sender, EventArgs e)
        {
            if (!_isInitialized || _populatingGrids || !Visible) return;
            var grid = sender as DataGridView;
            if (grid == null || grid != ActiveGrid) return;
            if (grid.SelectedRows.Count == 0) return;
            _currentIndex = grid.SelectedRows[0].Index;
            UpdateDetail();
        }

        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            var errors = ActiveErrors;
            if (e.RowIndex < 0 || e.RowIndex >= errors.Count) return;
            var error = errors[e.RowIndex];
            if (error.Range != null)
                DocumentHelper.GoToRange(error.Range);
        }

        private void BtnReplace_Click(object sender, EventArgs e)
        {
            var errors = ActiveErrors;
            if (_currentIndex < 0 || _currentIndex >= errors.Count) return;
            var error = errors[_currentIndex];

            string best = error.EnsureBestSuggestion();
            if (string.IsNullOrEmpty(best))
            {
                SafeExecutor.ShowWarning("Бу сўз учун таклиф топилмади.");
                return;
            }

            _onReplace?.Invoke(error, best);
            if (!error.IsResolved)
            {
                SafeExecutor.ShowWarning("Ҳужжат ёки матн ўзгарган. Қайта текширувни ишга туширинг.");
                return;
            }
            FilterErrors();
            ToastNotification.Success("Алмаштирилди", $"{error.Word} \u2192 {best}");
        }

        private void BtnIgnore_Click(object sender, EventArgs e)
        {
            var errors = ActiveErrors;
            if (_currentIndex < 0 || _currentIndex >= errors.Count) return;
            var error = errors[_currentIndex];

            _onIgnore?.Invoke(error);
            error.IsResolved = true;
            FilterErrors();
        }

        private void BtnAddDict_Click(object sender, EventArgs e)
        {
            if (IsGrammarTab) return; // Not applicable for grammar errors

            var errors = ActiveErrors;
            if (_currentIndex < 0 || _currentIndex >= errors.Count) return;
            var error = errors[_currentIndex];

            var result = _onAddToDict?.Invoke(error) ?? AddWordResult.Invalid;
            if (result == AddWordResult.Invalid) return;

            error.IsResolved = true;
            FilterErrors();
        }

        private void BtnGoTo_Click(object sender, EventArgs e)
        {
            var errors = ActiveErrors;
            if (_currentIndex < 0 || _currentIndex >= errors.Count) return;
            var error = errors[_currentIndex];
            if (error.Range != null)
                DocumentHelper.GoToRange(error.Range);
        }

        private void NavigateError(int direction)
        {
            var errors = ActiveErrors;
            var grid = ActiveGrid;
            if (errors.Count == 0) return;

            _currentIndex = (_currentIndex + direction + errors.Count) % errors.Count;
            grid.ClearSelection();
            grid.Rows[_currentIndex].Selected = true;
            grid.FirstDisplayedScrollingRowIndex = _currentIndex;

            var error = errors[_currentIndex];
            if (error.Range != null)
                DocumentHelper.GoToRange(error.Range);
        }

        private void UpdateDetail()
        {
            if (_detailLabel == null || _ruleLabel == null) return;

            var errors = ActiveErrors;
            if (errors == null || _currentIndex < 0 || _currentIndex >= errors.Count) return;
            var error = errors[_currentIndex];
            if (error == null) return;

            // Update "Add to Dict" visibility based on current tab
            _btnAddDict.Visible = !IsGrammarTab;

            // Only the selected entry performs lazy suggestion work, not every
            // row while the dialog is being constructed or filtered.
            error.EnsureBestSuggestion();

            if (error.IsGrammarError)
            {
                // Grammar error: show rule message in detail
                _detailLabel.Text = $"Сўз: {error.Word ?? "?"} \u2192 {error.BestSuggestion ?? "—"}    |    " +
                                   $"Ўрни: {error.ParagraphIndex}-параграф";
                _ruleLabel.Text = error.Message ?? error.Context ?? "";
                _ruleLabel.ForeColor = Color.FromArgb(56, 142, 60); // green tint for grammar
            }
            else
            {
                // Spelling error: show suggestions
                string suggestions = error.Suggestions != null && error.Suggestions.Count > 0
                    ? string.Join(", ", error.Suggestions.Select(s => s.Text))
                    : "таклиф йўқ";

                _detailLabel.Text = $"Сўз: {error.Word ?? "?"} \u2192 {error.BestSuggestion ?? "?"}    |    " +
                                   $"Ўрни: {error.ParagraphIndex}-параграф";
                _ruleLabel.Text = $"Вариантлар: {suggestions}";
                _ruleLabel.ForeColor = ThemeManager.TextSecondary;
            }

            // Update the grid cell with the now-loaded suggestion
            var grid = ActiveGrid;
            if (grid.SelectedRows.Count > 0 && error.BestSuggestion != null)
            {
                try { grid.SelectedRows[0].Cells["colSuggestion"].Value = error.BestSuggestion; }
                catch { }
            }
        }

        private static string TruncateContext(string text, int maxLen)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Length <= maxLen ? text : text.Substring(0, maxLen) + "...";
        }

        // =====================================================================
        //  PILL TAB BUTTONS
        // =====================================================================

        private Label CreatePill(string text, int tabIndex)
        {
            var pill = new Label
            {
                Text = text,
                Font = tabIndex == 0 ? _fPill : _fPillR,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                Tag = tabIndex
            };

            pill.Click += (s, e) => SelectPill(tabIndex);
            pill.Paint += PillPaint;
            return pill;
        }

        private void SelectPill(int tabIndex)
        {
            if (_activeTab == tabIndex) return;

            _activeTab = tabIndex;
            _currentIndex = -1;

            // Swap grid visibility
            _spellingGrid.Visible = !IsGrammarTab;
            _grammarGrid.Visible = IsGrammarTab;

            // Update pill styles
            _pillSpelling.Font = UiFont(_activeTab == 0 ? _fPill : _fPillR);
            _pillGrammar.Font = UiFont(_activeTab == 1 ? _fPill : _fPillR);
            _pillSpelling.Invalidate();
            _pillGrammar.Invalidate();

            UpdateStatusAndButtons();

            var grid = ActiveGrid;
            if (grid.Rows.Count > 0)
            {
                grid.ClearSelection();
                grid.Rows[0].Selected = true;
                _currentIndex = 0;
            }
            UpdateDetail();
        }

        private void PillPaint(object sender, PaintEventArgs e)
        {
            var pill = (Label)sender;
            int idx = (int)pill.Tag;
            bool active = idx == _activeTab;

            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = new Rectangle(0, 0, pill.Width - 1, pill.Height - 1);

            Color bg, fg, bdr;
            if (active)
            {
                // Spelling = red-tinted primary, Grammar = green-tinted
                if (idx == 0)
                {
                    bg = ThemeManager.Primary;
                    fg = ThemeManager.TextOnPrimary;
                    bdr = ThemeManager.Primary;
                }
                else
                {
                    bg = Color.FromArgb(56, 142, 60);
                    fg = Color.White;
                    bdr = Color.FromArgb(56, 142, 60);
                }
            }
            else
            {
                bg = ThemeManager.Surface;
                fg = ThemeManager.TextSecondary;
                bdr = ThemeManager.Border;
            }

            using (var path = RoundedRect(rect, 14))
            {
                using (var br = new SolidBrush(bg))
                    g.FillPath(br, path);
                using (var pen = new Pen(bdr))
                    g.DrawPath(pen, path);
            }

            TextRenderer.DrawText(g, pill.Text, pill.Font,
                new Rectangle(0, 0, pill.Width, pill.Height), fg,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        private static GraphicsPath RoundedRect(Rectangle rect, int r)
        {
            int d = r * 2;
            var gp = new GraphicsPath();
            gp.AddArc(rect.X, rect.Y, d, d, 180, 90);
            gp.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            gp.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            gp.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            gp.CloseFigure();
            return gp;
        }
    }
}



