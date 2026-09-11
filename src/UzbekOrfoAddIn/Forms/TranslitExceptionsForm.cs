using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Services;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>
    /// Fully modern transliteration exceptions manager.
    /// Design: 2025-era — pill category tabs, inline add row, glassmorphism cards,
    /// status chips, clean grid with hover effects, responsive layout.
    /// </summary>
    public class TranslitExceptionsForm : ModernForm
    {
        // === Callbacks ===
        private readonly Func<List<TranslitException>> _getExceptions;
        private readonly Action<TranslitException> _onAdd;
        private readonly Action<TranslitException> _onUpdate;
        private readonly Action<string> _onRemove;
        private readonly Action _onSave;

        // === Controls ===
        private DataGridView _grid;
        private ModernTextBox _searchBox;
        private ModernTextBox _txtOriginal;
        private ModernTextBox _txtReplacement;
        private ModernTextBox _txtDescription;
        private Label _countLabel;
        private Label _countTotal;
        private Panel _pillBar;
        private int _activePillIndex; // 0 = All
        private readonly List<Label> _pills = new List<Label>();

        // === Data ===
        private List<TranslitException> _allExceptions = new List<TranslitException>();
        private string _selectedCategoryKey; // null = all

        // Categories
        private static readonly string[] Categories =
            { "\u0411\u0440\u0435\u043d\u0434", "\u0422\u0435\u0445\u043d\u0438\u043a", "\u0413\u0435\u043e\u0433\u0440\u0430\u0444\u0438\u043a", "\u0418\u0441\u043c", "\u0411\u043e\u0448\u049b\u0430" };
        private static readonly string[] CategoryKeys =
            { "Brand", "Technical", "Geographic", "Name", "Other" };

        // Fonts (form-private, cached) — scaled up for readability
        private static readonly Font _fSM = new Font("Segoe UI", 10f, FontStyle.Regular);
        private static readonly Font _fBase = new Font("Segoe UI", 11f, FontStyle.Regular);
        private static readonly Font _fLG = new Font("Segoe UI", 12.5f, FontStyle.Regular);
        private static readonly Font _fLGBold = new Font("Segoe UI", 12.5f, FontStyle.Bold);
        private static readonly Font _fXL = new Font("Segoe UI", 16f, FontStyle.Bold);
        private static readonly Font _fPill = new Font("Segoe UI Semibold", 10.5f, FontStyle.Bold);
        private static readonly Font _fPillR = new Font("Segoe UI", 10.5f, FontStyle.Regular);

        // Colors for category chips
        private static readonly Color[] ChipColors =
        {
            Color.FromArgb(59, 130, 246),  // Brand  — blue
            Color.FromArgb(139, 92, 246),  // Tech   — violet
            Color.FromArgb(16, 185, 129),  // Geo    — emerald
            Color.FromArgb(245, 158, 11),  // Name   — amber
            Color.FromArgb(107, 114, 128), // Other  — gray
        };

        public TranslitExceptionsForm(
            Func<List<TranslitException>> getExceptions,
            Action<TranslitException> onAdd,
            Action<TranslitException> onUpdate,
            Action<string> onRemove,
            Action onSave)
        {
            _getExceptions = getExceptions;
            _onAdd = onAdd;
            _onUpdate = onUpdate;
            _onRemove = onRemove;
            _onSave = onSave;

            Title = "\u0418\u0441\u0442\u0438\u0441\u043d\u043e\u043b\u0430\u0440";
            Size = new Size(1200, 860);
            MinimumSize = new Size(920, 640);

            SuspendLayout();
            BuildUI();
            ResumeLayout(false);
            PerformLayout();

            Shown += (s, ev) => { RefreshData(); _searchBox.Focus(); };
        }

        // =================================================================
        //  HELPERS
        // =================================================================

        private static string CategoryToDisplay(string key)
        {
            if (string.IsNullOrEmpty(key)) return Categories[Categories.Length - 1];
            int idx = Array.FindIndex(CategoryKeys, k => k.Equals(key, StringComparison.OrdinalIgnoreCase));
            return idx >= 0 ? Categories[idx] : key;
        }

        private static string DisplayToCategory(string display)
        {
            if (string.IsNullOrEmpty(display)) return "Other";
            int idx = Array.IndexOf(Categories, display);
            return idx >= 0 ? CategoryKeys[idx] : display;
        }

        private static int CategoryKeyIndex(string key)
        {
            if (string.IsNullOrEmpty(key)) return CategoryKeys.Length - 1;
            int idx = Array.FindIndex(CategoryKeys, k => k.Equals(key, StringComparison.OrdinalIgnoreCase));
            return idx >= 0 ? idx : CategoryKeys.Length - 1;
        }

        // =================================================================
        //  BUILD UI
        // =================================================================

        private void BuildUI()
        {
            ContentPanel.Padding = new Padding(ThemeManager.SpaceLG, ThemeManager.SpaceSM,
                                               ThemeManager.SpaceLG, 0);

            // ─────────────────────────────────────────────────────────────
            //  SECTION 1: HEADER BAR — subtitle + stats
            // ─────────────────────────────────────────────────────────────
            var headerRow = new Panel { Dock = DockStyle.Top, Height = 32 };

            var lblSubtitle = new Label
            {
                Text = "\u0422\u0440\u0430\u043d\u0441\u043b\u0438\u0442\u0435\u0440\u0430\u0446\u0438\u044f\u0434\u0430 \u045e\u0437\u0433\u0430\u0440\u0442\u0438\u0440\u0438\u043b\u043c\u0430\u0439\u0434\u0438\u0433\u0430\u043d \u0441\u045e\u0437\u043b\u0430\u0440 \u0440\u045e\u0439\u0445\u0430\u0442\u0438",
                AutoSize = true,
                ForeColor = ThemeManager.TextSecondary,
                Font = _fBase,
                Location = new Point(0, 4)
            };

            _countTotal = new Label
            {
                Text = "",
                AutoSize = true,
                ForeColor = ThemeManager.TextSecondary,
                Font = _fBase,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                TextAlign = ContentAlignment.MiddleRight
            };
            _countTotal.Location = new Point(headerRow.Width - 200, 4);

            headerRow.Controls.AddRange(new Control[] { lblSubtitle, _countTotal });

            // ─────────────────────────────────────────────────────────────
            //  SECTION 2: SEARCH + PILL TABS + COUNT
            // ─────────────────────────────────────────────────────────────
            var toolbarRow = new Panel { Dock = DockStyle.Top, Height = 58 };

            _searchBox = new ModernTextBox
            {
                Placeholder = "\u049A\u0438\u0434\u0438\u0440\u0438\u0448...",
                Font = _fLG,
                Location = new Point(0, 6),
                Size = new Size(280, 42)
            };
            _searchBox.TextChanged += (s, e) => FilterData();

            // Pill category tabs
            _pillBar = new Panel
            {
                Location = new Point(294, 8),
                Size = new Size(580, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.Transparent
            };
            BuildPills();

            _countLabel = new Label
            {
                AutoSize = true,
                ForeColor = ThemeManager.TextSecondary,
                Font = _fSM,
                Visible = false        // filtered count merged into _countTotal
            };

            toolbarRow.Controls.AddRange(new Control[] { _searchBox, _pillBar });

            // ─────────────────────────────────────────────────────────────
            //  SECTION 3: INLINE ADD CARD (two rows: labels + inputs)
            // ─────────────────────────────────────────────────────────────
            var addSection = new Panel { Dock = DockStyle.Top, Height = 100, Padding = new Padding(0, 4, 0, 4) };

            var addCard = new InlineAddCard();
            addCard.Dock = DockStyle.Fill;
            addCard.Height = 92;

            // --- Label/Input geometry ---
            int lblY = 4;
            int inputY = 36;
            int fieldW = 180;
            int fieldH = 40;
            int origX = 12;
            int arrowGap = 8;
            int arrowSlotW = 30; // dedicated space so arrow never touches fields on high DPI
            int replX = origX + fieldW + arrowGap + arrowSlotW + arrowGap;
            int descX = replX + fieldW + 10;

            // --- Label row ---
            var lblOrig = new Label
            {
                Text = "\u041b\u043e\u0442\u0438\u043d \u0448\u0430\u043a\u043b\u0438",
                AutoSize = true,
                Font = _fSM,
                ForeColor = ThemeManager.TextSecondary,
                BackColor = ThemeManager.Surface,
                Location = new Point(origX + 2, lblY)
            };
            var lblRepl = new Label
            {
                Text = "\u041a\u0438\u0440\u0438\u043b\u043b \u0448\u0430\u043a\u043b\u0438",
                AutoSize = true,
                Font = _fSM,
                ForeColor = ThemeManager.TextSecondary,
                BackColor = ThemeManager.Surface,
                Location = new Point(replX + 2, lblY)
            };
            var lblDescAdd = new Label
            {
                Text = "\u0422\u0430\u0432\u0441\u0438\u0444\u0438 (\u0438\u0445\u0442\u0438\u044f\u0440\u0438\u0439)",
                AutoSize = true,
                Font = _fSM,
                ForeColor = ThemeManager.TextSecondary,
                BackColor = ThemeManager.Surface,
                Location = new Point(descX + 2, lblY)
            };

            // --- Input row (labels ~20px tall, start inputs below with gap) ---

            _txtOriginal = new ModernTextBox
            {
                Font = _fBase,
                Placeholder = "\u043c\u0430\u0441: Google",
                Size = new Size(fieldW, fieldH),
                Location = new Point(origX, inputY)
            };

            var arrowLabel = new Label
            {
                Text = "\u2192",
                AutoSize = false,
                Font = _fLG,
                ForeColor = ThemeManager.TextSecondary,
                BackColor = ThemeManager.Surface,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(arrowSlotW, fieldH),
                Location = new Point(origX + fieldW + arrowGap, inputY)
            };

            _txtReplacement = new ModernTextBox
            {
                Font = _fBase,
                Placeholder = "\u043c\u0430\u0441: \u0413\u0443\u0433\u043b",
                Size = new Size(fieldW, fieldH),
                Location = new Point(replX, inputY)
            };

            _txtDescription = new ModernTextBox
            {
                Font = _fBase,
                Placeholder = "\u049a\u0438\u0441\u049b\u0430\u0447\u0430 \u0438\u0437\u043e\u04b3...",
                Size = new Size(380, 40),
                Location = new Point(descX, inputY),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var btnAdd = new ModernButton
            {
                Text = "+ \u049A\u045e\u0448\u0438\u0448",
                Style = ModernButton.ButtonStyle.Primary,
                Font = _fBase,
                Size = new Size(120, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Location = new Point(800, inputY)
            };
            btnAdd.Click += BtnAdd_Click;

            // Enter key in text boxes triggers add
            _txtOriginal.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { BtnAdd_Click(s, e); e.Handled = true; } };
            _txtReplacement.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { BtnAdd_Click(s, e); e.Handled = true; } };
            _txtDescription.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { BtnAdd_Click(s, e); e.Handled = true; } };

            addCard.Controls.AddRange(new Control[]
            {
                lblOrig, lblRepl, lblDescAdd,
                _txtOriginal, arrowLabel, _txtReplacement, _txtDescription, btnAdd
            });

            // Position description and button correctly when card gets real width
            addCard.Layout += (s, eL) =>
            {
                int cW = addCard.ClientSize.Width;
                int addBtnW = Px(120);
                int gap = 10;
                int descRight = cW - addBtnW - Px(gap) - Px(12); // 12 right padding
                int descLeft = Px(descX);
                if (descRight > descLeft)
                    _txtDescription.Width = descRight - descLeft;
                btnAdd.Location = new Point(cW - addBtnW - Px(12), Px(inputY));
            };

            addSection.Controls.Add(addCard);

            // ─────────────────────────────────────────────────────────────
            //  SECTION 4: DATA GRID
            // ─────────────────────────────────────────────────────────────
            BuildGrid();

            // ─────────────────────────────────────────────────────────────
            //  SECTION 5: ACTION BAR
            // ─────────────────────────────────────────────────────────────
            BuildActionBar();

            // ─────────────────────────────────────────────────────────────
            //  ASSEMBLE (bottom-up for DockStyle.Top)
            // ─────────────────────────────────────────────────────────────
            ContentPanel.Controls.Add(_grid);
            ContentPanel.Controls.Add(addSection);
            ContentPanel.Controls.Add(toolbarRow);
            ContentPanel.Controls.Add(headerRow);
        }

        // =================================================================
        //  PILL CATEGORY TABS
        // =================================================================

        private void BuildPills()
        {
            _pills.Clear();
            _pillBar.Controls.Clear();

            string[] labels = new string[Categories.Length + 1];
            labels[0] = "\u04B2\u0430\u043c\u043c\u0430\u0441\u0438";
            for (int i = 0; i < Categories.Length; i++) labels[i + 1] = Categories[i];

            int x = 0;
            for (int i = 0; i < labels.Length; i++)
            {
                var pill = new Label
                {
                    Text = labels[i],
                    Font = i == 0 ? _fPill : _fPillR,
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    Tag = i
                };

                // Measure text width + padding
                int tw = TextRenderer.MeasureText(labels[i], _fPill).Width + 28;
                pill.Size = new Size(tw, 36);
                pill.Location = new Point(x, 2);

                int idx = i;
                pill.Click += (s, e) => SelectPill(idx);
                pill.Paint += PillPaint;

                _pills.Add(pill);
                _pillBar.Controls.Add(pill);
                x += tw + 8;
            }

            _activePillIndex = 0;
            _selectedCategoryKey = null;
            UpdatePillStyles();
        }

        private void SelectPill(int index)
        {
            _activePillIndex = index;
            _selectedCategoryKey = index == 0 ? null : CategoryKeys[index - 1];
            UpdatePillStyles();
            FilterData();
        }

        private void UpdatePillStyles()
        {
            for (int i = 0; i < _pills.Count; i++)
            {
                _pills[i].Font = UiFont(i == _activePillIndex ? _fPill : _fPillR);
                _pills[i].Invalidate();
            }
        }

        private void PillPaint(object sender, PaintEventArgs e)
        {
            var pill = (Label)sender;
            int idx = (int)pill.Tag;
            bool active = idx == _activePillIndex;

            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = new Rectangle(0, 0, pill.Width - 1, pill.Height - 1);

            Color bg, fg, bdr;
            if (active)
            {
                bg = ThemeManager.Primary;
                fg = ThemeManager.TextOnPrimary;
                bdr = ThemeManager.Primary;
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

        // =================================================================
        //  DATA GRID
        // =================================================================

        private void BuildGrid()
        {
            _grid = new DpiDataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                BackgroundColor = ThemeManager.Background,
                GridColor = ThemeManager.IsDarkTheme
                    ? Color.FromArgb(50, 50, 50) : Color.FromArgb(240, 240, 240),
                EnableHeadersVisualStyles = false,
                RowTemplate = { Height = 48 },
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                ScrollBars = ScrollBars.Vertical
            };

            // Columns
            var colEnabled = new DataGridViewCheckBoxColumn
            {
                Name = "colEnabled",
                HeaderText = "Фаол",
                Width = 75,
                FillWeight = 3,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            };

            var colOriginal = new DataGridViewTextBoxColumn
            {
                Name = "colOriginal",
                HeaderText = "\u041b\u043e\u0442\u0438\u043d",
                FillWeight = 18,
                ReadOnly = true
            };

            var colReplacement = new DataGridViewTextBoxColumn
            {
                Name = "colReplacement",
                HeaderText = "\u041a\u0438\u0440\u0438\u043b\u043b",
                FillWeight = 18,
                ReadOnly = true
            };

            var colCategory = new DataGridViewTextBoxColumn
            {
                Name = "colCategory",
                HeaderText = "\u041a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f",
                FillWeight = 12,
                ReadOnly = true
            };

            var colDescription = new DataGridViewTextBoxColumn
            {
                Name = "colDescription",
                HeaderText = "\u0422\u0430\u0432\u0441\u0438\u0444\u0438",
                FillWeight = 28,
                ReadOnly = true
            };

            _grid.Columns.AddRange(new DataGridViewColumn[]
            {
                colEnabled, colOriginal, colReplacement, colCategory, colDescription
            });

            // Header style
            _grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = ThemeManager.Background,
                ForeColor = ThemeManager.TextSecondary,
                Font = _fSM,
                SelectionBackColor = ThemeManager.Background,
                SelectionForeColor = ThemeManager.TextSecondary,
                Padding = new Padding(8, 0, 0, 0),
                Alignment = DataGridViewContentAlignment.MiddleLeft
            };
            _grid.ColumnHeadersHeight = 40;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // Cell style
            _grid.DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = ThemeManager.Background,
                ForeColor = ThemeManager.TextPrimary,
                Font = _fBase,
                SelectionBackColor = ThemeManager.SelectedRow,
                SelectionForeColor = ThemeManager.TextPrimary,
                Padding = new Padding(8, 0, 0, 0)
            };

            _grid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = ThemeManager.IsDarkTheme
                    ? Color.FromArgb(36, 36, 36) : Color.FromArgb(250, 251, 252),
                ForeColor = ThemeManager.TextPrimary,
                SelectionBackColor = ThemeManager.SelectedRow,
                SelectionForeColor = ThemeManager.TextPrimary,
                Padding = new Padding(8, 0, 0, 0)
            };

            // Custom paint for category chips
            _grid.CellPainting += Grid_CellPainting;

            // Events
            _grid.CellDoubleClick += Grid_CellDoubleClick;
            _grid.KeyDown += Grid_KeyDown;
            _grid.CellValueChanged += Grid_CellValueChanged;
            _grid.CurrentCellDirtyStateChanged += Grid_CurrentCellDirtyStateChanged;
        }

        // =================================================================
        //  ACTION BAR
        // =================================================================

        private void BuildActionBar()
        {
            ActionBar.Height = 68;
            ActionBar.BackColor = ThemeManager.Background;

            // Separator painted at top of ActionBar
            ActionBar.Paint += (s, pe) =>
            {
                using (var pen = new Pen(ThemeManager.Border))
                    pe.Graphics.DrawLine(pen, 0, 0, ActionBar.Width, 0);
            };

            int btnH = 42;
            int btnY = (ActionBar.Height - btnH) / 2;

            var btnDelete = new ModernButton
            {
                Text = "\u040E\u0447\u0438\u0440\u0438\u0448",
                Style = ModernButton.ButtonStyle.Danger,
                Font = _fBase,
                Size = new Size(116, btnH),
                Location = new Point(ThemeManager.SpaceLG, btnY)
            };
            btnDelete.Click += BtnDelete_Click;

            var btnExport = new ModernButton
            {
                Text = "\u042D\u043a\u0441\u043f\u043e\u0440\u0442",
                Style = ModernButton.ButtonStyle.Ghost,
                Font = _fBase,
                Size = new Size(116, btnH),
                Location = new Point(ThemeManager.SpaceLG + 116 + 10, btnY)
            };
            btnExport.Click += BtnExport_Click;

            var btnImport = new ModernButton
            {
                Text = "\u0418\u043c\u043f\u043e\u0440\u0442",
                Style = ModernButton.ButtonStyle.Ghost,
                Font = _fBase,
                Size = new Size(116, btnH),
                Location = new Point(ThemeManager.SpaceLG + (116 + 10) * 2, btnY)
            };
            btnImport.Click += BtnImport_Click;

            var btnClose = new ModernButton
            {
                Text = "\u0401\u043f\u0438\u0448",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = _fBase,
                Size = new Size(110, btnH),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnClose.Location = new Point(ActionBar.Width - 110 - ThemeManager.SpaceLG, btnY);
            btnClose.Click += (s, e) => Close();

            var btnSave = new ModernButton
            {
                Text = "\u0421\u0430\u049b\u043b\u0430\u0448",
                Style = ModernButton.ButtonStyle.Primary,
                Font = _fBase,
                Size = new Size(130, btnH),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnSave.Location = new Point(ActionBar.Width - 110 - 130 - ThemeManager.SpaceLG - 10, btnY);
            btnSave.Click += (s, e) =>
            {
                _onSave?.Invoke();
                ToastNotification.Success("\u0421\u0430\u049b\u043b\u0430\u043d\u0434\u0438",
                    $"{_allExceptions.Count} \u0442\u0430 \u0438\u0441\u0442\u0438\u0441\u043d\u043e \u0441\u0430\u049b\u043b\u0430\u043d\u0434\u0438.");
            };

            ActionBar.Controls.AddRange(new Control[] { btnDelete, btnExport, btnImport, btnSave, btnClose });
        }

        // =================================================================
        //  CATEGORY CHIP PAINTING IN GRID
        // =================================================================

        private void Grid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (e.ColumnIndex < 0) return;
            if (_grid.Columns[e.ColumnIndex].Name != "colCategory") return;

            e.Handled = true;

            // Paint background
            using (var bg = new SolidBrush(e.CellStyle.BackColor))
                e.Graphics.FillRectangle(bg, e.CellBounds);

            // Paint bottom border
            using (var pen = new Pen(_grid.GridColor))
                e.Graphics.DrawLine(pen, e.CellBounds.Left, e.CellBounds.Bottom - 1,
                    e.CellBounds.Right, e.CellBounds.Bottom - 1);

            // Paint selection overlay
            if ((e.State & DataGridViewElementStates.Selected) != 0)
            {
                using (var sel = new SolidBrush(Color.FromArgb(40, ThemeManager.Primary)))
                    e.Graphics.FillRectangle(sel, e.CellBounds);
            }

            string text = e.Value?.ToString() ?? "";
            if (string.IsNullOrEmpty(text)) return;

            // Determine chip color
            int catIdx = Array.IndexOf(Categories, text);
            Color chipColor = catIdx >= 0 && catIdx < ChipColors.Length
                ? ChipColors[catIdx] : ChipColors[ChipColors.Length - 1];

            // Check if row is disabled
            var row = _grid.Rows[e.RowIndex];
            var exc = row.Tag as TranslitException;
            bool dimmed = exc != null && !exc.Enabled;

            if (dimmed) chipColor = Color.FromArgb(80, chipColor);

            // Measure chip
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var chipSize = TextRenderer.MeasureText(text, UiFont(_fSM));
            int chipW = chipSize.Width + Px(16);
            int chipH = Px(22);
            int chipX = e.CellBounds.X + Px(8);
            int chipY = e.CellBounds.Y + (e.CellBounds.Height - chipH) / 2;

            var chipRect = new Rectangle(chipX, chipY, chipW, chipH);

            // Draw chip background (subtle tint)
            Color chipBg = Color.FromArgb(ThemeManager.IsDarkTheme ? 35 : 20, chipColor);
            using (var path = RoundedRect(chipRect, Px(11)))
            using (var br = new SolidBrush(chipBg))
                g.FillPath(br, path);

            // Draw chip text
            Color chipFg = dimmed ? ThemeManager.TextDisabled : chipColor;
            TextRenderer.DrawText(g, text, UiFont(_fSM), chipRect, chipFg,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        // =================================================================
        //  DATA
        // =================================================================

        private void RefreshData()
        {
            _allExceptions = _getExceptions?.Invoke() ?? new List<TranslitException>();
            UpdateTotalCount();
            FilterData();
        }

        private void UpdateTotalCount()
        {
            int enabled = _allExceptions.Count(e => e.Enabled);
            _countTotal.Text = $"{_allExceptions.Count} \u0442\u0430 \u0438\u0441\u0442\u0438\u0441\u043d\u043e  \u00b7  {enabled} \u0444\u0430\u043e\u043b";
            _countTotal.Location = new Point(
                _countTotal.Parent.Width - _countTotal.Width - 4, _countTotal.Location.Y);
        }

        /// <summary>
        /// Updates the header stats when a filter reduces the visible set.
        /// </summary>
        private void UpdateHeaderStats(int visibleCount)
        {
            int total = _allExceptions.Count;
            int enabled = _allExceptions.Count(e => e.Enabled);
            if (visibleCount < total)
                _countTotal.Text = $"{visibleCount} / {total} \u0442\u0430 \u0438\u0441\u0442\u0438\u0441\u043d\u043e  \u00b7  {enabled} \u0444\u0430\u043e\u043b";
            else
                _countTotal.Text = $"{total} \u0442\u0430 \u0438\u0441\u0442\u0438\u0441\u043d\u043e  \u00b7  {enabled} \u0444\u0430\u043e\u043b";
            _countTotal.Location = new Point(
                _countTotal.Parent.Width - _countTotal.Width - 4, _countTotal.Location.Y);
        }

        private void FilterData()
        {
            string query = _searchBox?.Text?.Trim().ToLowerInvariant() ?? "";

            var filtered = _allExceptions.AsEnumerable();

            if (!string.IsNullOrEmpty(query))
            {
                filtered = filtered.Where(e =>
                    (e.Original?.ToLowerInvariant().Contains(query) ?? false) ||
                    (e.Replacement?.ToLowerInvariant().Contains(query) ?? false) ||
                    (e.Description?.ToLowerInvariant().Contains(query) ?? false));
            }

            if (_selectedCategoryKey != null)
            {
                filtered = filtered.Where(e =>
                    (e.Category ?? "Other").Equals(_selectedCategoryKey, StringComparison.OrdinalIgnoreCase));
            }

            var list = filtered.ToList();

            _grid.SuspendLayout();
            _grid.Rows.Clear();

            if (list.Count > 0)
            {
                var rows = new DataGridViewRow[list.Count];
                for (int i = 0; i < list.Count; i++)
                {
                    var exc = list[i];
                    var row = new DataGridViewRow { Height = Px(44) };
                    row.CreateCells(_grid,
                        exc.Enabled,
                        exc.Original ?? "",
                        exc.Replacement ?? "",
                        CategoryToDisplay(exc.Category),
                        exc.Description ?? "");
                    row.Tag = exc;

                    if (!exc.Enabled)
                    {
                        for (int c = 1; c < row.Cells.Count; c++)
                        {
                            row.Cells[c].Style.ForeColor = ThemeManager.TextDisabled;
                            row.Cells[c].Style.SelectionForeColor = ThemeManager.TextDisabled;
                        }
                    }

                    rows[i] = row;
                }
                _grid.Rows.AddRange(rows);
            }

            _countLabel.Text = $"{list.Count} \u043a\u045e\u0440\u0441\u0430\u0442\u0438\u043b\u043c\u043e\u049b\u0434\u0430";
            // Update header with filtered count when searching/filtering
            UpdateHeaderStats(list.Count);
            _grid.ResumeLayout();
        }

        // =================================================================
        //  EVENTS
        // =================================================================

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string original = _txtOriginal?.Text?.Trim();
            string replacement = _txtReplacement?.Text?.Trim();
            string description = _txtDescription?.Text?.Trim();

            if (string.IsNullOrEmpty(original))
            {
                SafeExecutor.ShowWarning("\u041b\u043e\u0442\u0438\u043d\u0447\u0430 \u0441\u045e\u0437\u043d\u0438 \u043a\u0438\u0440\u0438\u0442\u0438\u043d\u0433.");
                return;
            }
            if (string.IsNullOrEmpty(replacement))
            {
                SafeExecutor.ShowWarning("\u041a\u0438\u0440\u0438\u043b\u043b\u0447\u0430 \u0442\u0430\u0440\u0436\u0438\u043c\u0430\u043d\u0438 \u043a\u0438\u0440\u0438\u0442\u0438\u043d\u0433.");
                return;
            }

            // Auto-detect category from active pill (if not "All")
            string category = _selectedCategoryKey ?? "Other";

            _onAdd?.Invoke(new TranslitException(original, replacement, category, description));

            _txtOriginal.Text = "";
            _txtReplacement.Text = "";
            _txtDescription.Text = "";
            _txtOriginal.Focus();

            RefreshData();
            ToastNotification.Success("\u049A\u045e\u0448\u0438\u043b\u0434\u0438", $"{original} \u2192 {replacement}");
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_grid.SelectedRows.Count == 0)
            {
                SafeExecutor.ShowWarning("\u040E\u0447\u0438\u0440\u0438\u043b\u0430\u0434\u0438\u0433\u0430\u043d \u0441\u0430\u0442\u0440\u043d\u0438 \u0442\u0430\u043d\u043b\u0430\u043d\u0433.");
                return;
            }

            var row = _grid.SelectedRows[0];
            string original = row.Cells["colOriginal"].Value?.ToString();
            if (string.IsNullOrEmpty(original)) return;

            if (SafeExecutor.Confirm($"\u00ab{original}\u00bb \u0438\u0441\u0442\u0438\u0441\u043d\u043e\u0441\u0438\u043d\u0438 \u045e\u0447\u0438\u0440\u043c\u043e\u049b\u0447\u0438\u043c\u0438\u0441\u0438\u0437?"))
            {
                _onRemove?.Invoke(original);
                RefreshData();
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            using (var dlg = new SaveFileDialog
            {
                Filter = "CSV (*.csv)|*.csv|JSON (*.json)|*.json|Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls",
                FileName = "translit_exceptions",
                DefaultExt = "csv"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                try
                {
                    string ext = Path.GetExtension(dlg.FileName)?.ToLowerInvariant();
                    switch (ext)
                    {
                        case ".json":
                            ExportJson(dlg.FileName);
                            break;
                        case ".xlsx":
                            ExportXlsx(dlg.FileName);
                            break;
                        case ".xls":
                            ExportXls(dlg.FileName);
                            break;
                        default:
                            ExportCsv(dlg.FileName);
                            break;
                    }

                    ToastNotification.Success("\u042D\u043a\u0441\u043f\u043e\u0440\u0442",
                        $"{_allExceptions.Count} \u0442\u0430 \u0438\u0441\u0442\u0438\u0441\u043d\u043e \u044d\u043a\u0441\u043f\u043e\u0440\u0442 \u049b\u0438\u043b\u0438\u043d\u0434\u0438.");
                }
                catch (Exception ex)
                {
                    SafeExecutor.ShowError($"\u042d\u043a\u0441\u043f\u043e\u0440\u0442\u0434\u0430 \u0445\u0430\u0442\u043e\u043b\u0438\u043a: {ex.Message}");
                }
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog
            {
                Filter = "CSV (*.csv)|*.csv|JSON (*.json)|*.json|Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls",
                FilterIndex = 1,
                Title = "\u0418\u0441\u0442\u0438\u0441\u043d\u043e\u043b\u0430\u0440\u043d\u0438 \u0438\u043c\u043f\u043e\u0440\u0442 \u049b\u0438\u043b\u0438\u0448"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                try
                {
                    List<TranslitException> imported;
                    string ext = Path.GetExtension(dlg.FileName)?.ToLowerInvariant();
                    switch (ext)
                    {
                        case ".json":
                            imported = ParseImportedJson(dlg.FileName);
                            break;
                        case ".xlsx":
                            imported = ParseImportedXlsx(dlg.FileName);
                            break;
                        case ".xls":
                            imported = ParseImportedXls(dlg.FileName);
                            break;
                        default:
                            imported = ParseImportedCsv(dlg.FileName);
                            break;
                    }

                    if (imported == null || imported.Count == 0)
                    {
                        ToastNotification.ShowInfo("\u0418\u043c\u043f\u043e\u0440\u0442", "\u0418\u043c\u043f\u043e\u0440\u0442 \u0443\u0447\u0443\u043d \u043c\u0430\u044a\u043b\u0443\u043c\u043e\u0442 \u0442\u043e\u043f\u0438\u043b\u043c\u0430\u0434\u0438.");
                        return;
                    }

                    var knownOriginals = new HashSet<string>(
                        _allExceptions
                            .Where(x => !string.IsNullOrWhiteSpace(x?.Original))
                            .Select(x => x.Original.Trim()),
                        StringComparer.OrdinalIgnoreCase);

                    int added = 0;
                    int updated = 0;
                    int skipped = 0;

                    foreach (var item in imported)
                    {
                        var normalized = NormalizeImportedException(item);
                        if (normalized == null)
                        {
                            skipped++;
                            continue;
                        }

                        if (knownOriginals.Contains(normalized.Original))
                            updated++;
                        else
                        {
                            added++;
                            knownOriginals.Add(normalized.Original);
                        }

                        _onAdd?.Invoke(normalized);
                    }

                    if (added == 0 && updated == 0)
                    {
                        ToastNotification.ShowInfo("\u0418\u043c\u043f\u043e\u0440\u0442", "\u0418\u043c\u043f\u043e\u0440\u0442\u0434\u0430 \u049b\u045e\u0448\u0438\u0448 \u0443\u0447\u0443\u043d \u044f\u0440\u043e\u049b\u043b\u0438 \u0441\u0430\u0442\u0440 \u0442\u043e\u043f\u0438\u043b\u043c\u0430\u0434\u0438.");
                        return;
                    }

                    _onSave?.Invoke();
                    RefreshData();

                    string msg = $"{added} \u0442\u0430 \u049b\u045e\u0448\u0438\u043b\u0434\u0438, {updated} \u0442\u0430 \u044f\u043d\u0433\u0438\u043b\u0430\u043d\u0434\u0438.";
                    if (skipped > 0) msg += $" {skipped} \u0442\u0430 \u0441\u0430\u0442\u0440 \u045e\u0442\u043a\u0430\u0437\u0438\u0431 \u044e\u0431\u043e\u0440\u0438\u043b\u0434\u0438.";

                    ToastNotification.Success("\u0418\u043c\u043f\u043e\u0440\u0442 \u0442\u0443\u0433\u0430\u0434\u0438", msg);
                }
                catch (Exception ex)
                {
                    SafeExecutor.ShowError($"\u0418\u043c\u043f\u043e\u0440\u0442\u0434\u0430 \u0445\u0430\u0442\u043e\u043b\u0438\u043a: {ex.Message}");
                }
            }
        }

        // =================================================================
        //  GRID EVENTS
        // =================================================================

        private void Grid_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (_grid.IsCurrentCellDirty && _grid.CurrentCell is DataGridViewCheckBoxCell)
                _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (_grid.Columns[e.ColumnIndex].Name != "colEnabled") return;

            var row = _grid.Rows[e.RowIndex];
            var exc = row.Tag as TranslitException;
            if (exc == null) return;

            bool enabled = row.Cells["colEnabled"].Value is bool b && b;
            exc.Enabled = enabled;
            _onUpdate?.Invoke(exc);

            var color = enabled ? ThemeManager.TextPrimary : ThemeManager.TextDisabled;
            for (int c = 1; c < row.Cells.Count; c++)
            {
                row.Cells[c].Style.ForeColor = color;
                row.Cells[c].Style.SelectionForeColor = color;
            }
            _grid.InvalidateRow(e.RowIndex);
            UpdateTotalCount();
        }

        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row = _grid.Rows[e.RowIndex];
            _txtOriginal.Text = row.Cells["colOriginal"].Value?.ToString() ?? "";
            _txtReplacement.Text = row.Cells["colReplacement"].Value?.ToString() ?? "";
            _txtDescription.Text = row.Cells["colDescription"].Value?.ToString() ?? "";
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                BtnDelete_Click(sender, e);
                e.Handled = true;
            }
        }

        // =================================================================
        //  EXPORT
        // =================================================================

        private void ExportCsv(string path)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Enabled,Original,Replacement,Category,Description");
            foreach (var exc in _allExceptions)
            {
                sb.AppendLine(string.Join(",",
                    exc.Enabled ? "true" : "false",
                    CsvEsc(exc.Original),
                    CsvEsc(exc.Replacement),
                    CsvEsc(exc.Category ?? "Other"),
                    CsvEsc(exc.Description ?? "")));
            }
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        private void ExportJson(string path)
        {
            var sb = new StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < _allExceptions.Count; i++)
            {
                var ex = _allExceptions[i];
                sb.Append("  { ");
                sb.Append($"\"Original\": \"{JsonEsc(ex.Original)}\", ");
                sb.Append($"\"Replacement\": \"{JsonEsc(ex.Replacement)}\", ");
                sb.Append($"\"Category\": \"{JsonEsc(ex.Category ?? "Other")}\", ");
                sb.Append($"\"Description\": \"{JsonEsc(ex.Description ?? "")}\", ");
                sb.Append($"\"Enabled\": {(ex.Enabled ? "true" : "false")}");
                sb.Append(" }");
                if (i < _allExceptions.Count - 1) sb.Append(",");
                sb.AppendLine();
            }
            sb.Append("]");
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        private void ExportXlsx(string path)
        {
            if (File.Exists(path))
                File.Delete(path);

            using (var fs = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create, false))
            {
                WriteZipEntry(zip, "[Content_Types].xml",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
  <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
  <Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/>
  <Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml""/>
</Types>");

                WriteZipEntry(zip, "_rels/.rels",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/>
  <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml""/>
</Relationships>");

                WriteZipEntry(zip, "docProps/app.xml",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties""
            xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">
  <Application>UzbekOrfo</Application>
</Properties>");

                WriteZipEntry(zip, "docProps/core.xml",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties""
                   xmlns:dc=""http://purl.org/dc/elements/1.1/""
                   xmlns:dcterms=""http://purl.org/dc/terms/""
                   xmlns:dcmitype=""http://purl.org/dc/dcmitype/""
                   xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <dc:creator>UzbekOrfo</dc:creator>
  <cp:lastModifiedBy>UzbekOrfo</cp:lastModifiedBy>
  <dcterms:created xsi:type=""dcterms:W3CDTF"">" + DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") + @"</dcterms:created>
  <dcterms:modified xsi:type=""dcterms:W3CDTF"">" + DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") + @"</dcterms:modified>
</cp:coreProperties>");

                WriteZipEntry(zip, "xl/workbook.xml",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
          xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <sheets>
    <sheet name=""Exceptions"" sheetId=""1"" r:id=""rId1""/>
  </sheets>
</workbook>");

                WriteZipEntry(zip, "xl/_rels/workbook.xml.rels",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>
</Relationships>");

                WriteZipEntry(zip, "xl/styles.xml",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <fonts count=""1""><font><sz val=""11""/><name val=""Calibri""/></font></fonts>
  <fills count=""1""><fill><patternFill patternType=""none""/></fill></fills>
  <borders count=""1""><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>
  <cellXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/></cellXfs>
  <cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>
</styleSheet>");

                WriteZipEntry(zip, "xl/worksheets/sheet1.xml", BuildExceptionsSheetXml(_allExceptions));
            }
        }

        private void ExportXls(string path)
        {
            string tempXlsx = Path.Combine(
                Path.GetTempPath(),
                "uzbekorfo_exceptions_" + Guid.NewGuid().ToString("N") + ".xlsx");

            try
            {
                ExportXlsx(tempXlsx);
                ConvertXlsxToXlsWithExcel(tempXlsx, path);
            }
            finally
            {
                try
                {
                    if (File.Exists(tempXlsx))
                        File.Delete(tempXlsx);
                }
                catch { }
            }
        }

        private static void ConvertXlsxToXlsWithExcel(string sourceXlsx, string targetXls)
        {
            object excelApp = null;
            object workbooks = null;
            object workbook = null;

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                    throw new InvalidOperationException("Excel is not installed for XLS export.");

                excelApp = Activator.CreateInstance(excelType);
                dynamic app = excelApp;
                app.Visible = false;
                app.DisplayAlerts = false;

                workbooks = app.Workbooks;
                dynamic books = workbooks;
                workbook = books.Open(sourceXlsx);
                dynamic wb = workbook;

                // 56 = xlExcel8 (.xls)
                wb.SaveAs(targetXls, 56);
                wb.Close(false);
                app.Quit();
            }
            finally
            {
                try
                {
                    if (workbook != null)
                        ((dynamic)workbook).Close(false);
                }
                catch { }

                try
                {
                    if (excelApp != null)
                        ((dynamic)excelApp).Quit();
                }
                catch { }

                ReleaseComObject(workbook);
                ReleaseComObject(workbooks);
                ReleaseComObject(excelApp);
            }
        }

        private static string BuildExceptionsSheetXml(List<TranslitException> exceptions)
        {
            if (exceptions == null) exceptions = new List<TranslitException>();

            var sb = new StringBuilder(8192);
            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
            sb.AppendLine(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            sb.AppendLine(@"  <sheetViews><sheetView workbookViewId=""0""/></sheetViews>");
            sb.AppendLine(@"  <sheetFormatPr defaultRowHeight=""15""/>");
            sb.AppendLine(@"  <cols>");
            sb.AppendLine(@"    <col min=""1"" max=""1"" width=""10"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""2"" max=""2"" width=""24"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""3"" max=""3"" width=""24"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""4"" max=""4"" width=""18"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""5"" max=""5"" width=""48"" customWidth=""1""/>");
            sb.AppendLine(@"  </cols>");
            sb.AppendLine(@"  <sheetData>");

            AppendInlineRow(sb, 1, new[] { "Enabled", "Original", "Replacement", "Category", "Description" });

            int row = 2;
            for (int i = 0; i < exceptions.Count; i++)
            {
                var exc = exceptions[i] ?? new TranslitException();
                AppendInlineRow(sb, row++, new[]
                {
                    exc.Enabled ? "true" : "false",
                    exc.Original ?? string.Empty,
                    exc.Replacement ?? string.Empty,
                    exc.Category ?? "Other",
                    exc.Description ?? string.Empty
                });
            }

            sb.AppendLine(@"  </sheetData>");
            sb.AppendLine(@"</worksheet>");
            return sb.ToString();
        }

        private static void WriteZipEntry(ZipArchive zip, string entryPath, string content)
        {
            var entry = zip.CreateEntry(entryPath, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(content);
            }
        }

        private static void AppendInlineRow(StringBuilder sb, int rowIndex, string[] values)
        {
            sb.Append("    <row r=\"").Append(rowIndex).AppendLine("\">");
            for (int i = 0; i < values.Length; i++)
            {
                string cellRef = GetExcelColumnName(i + 1) + rowIndex;
                string value = EscapeXmlForXlsx(values[i] ?? string.Empty);
                sb.Append("      <c r=\"").Append(cellRef).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">")
                  .Append(value)
                  .AppendLine("</t></is></c>");
            }
            sb.AppendLine("    </row>");
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            var sb = new StringBuilder();
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                sb.Insert(0, (char)('A' + modulo));
                columnNumber = (columnNumber - modulo) / 26;
            }
            return sb.ToString();
        }

        private static string EscapeXmlForXlsx(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;

            var sb = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++)
            {
                char ch = value[i];
                if (!IsValidXmlChar(ch)) continue;

                switch (ch)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    case '\'': sb.Append("&apos;"); break;
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }

        private static bool IsValidXmlChar(int c)
        {
            return c == 0x9 || c == 0xA || c == 0xD ||
                   (c >= 0x20 && c <= 0xD7FF) ||
                   (c >= 0xE000 && c <= 0xFFFD);
        }

        private static List<TranslitException> ParseImportedJson(string path)
        {
            string json = File.ReadAllText(path, Encoding.UTF8);
            return TransliterationService.SimpleJsonParser.ParseExceptions(json)
                   ?? new List<TranslitException>();
        }

        private static List<TranslitException> ParseImportedCsv(string path)
        {
            string csv = File.ReadAllText(path, Encoding.UTF8);
            var rows = ParseCsvRows(csv);
            var table = rows
                .Select(r => (r ?? new List<string>()).ToArray())
                .ToList();
            return ParseExceptionsFromTable(table);
        }

        private static List<TranslitException> ParseImportedXls(string path)
        {
            var table = new List<string[]>();

            object excelApp = null;
            object workbooks = null;
            object workbook = null;
            object worksheets = null;
            object worksheet = null;
            object usedRange = null;

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                    throw new InvalidOperationException("Excel is not installed for XLS import.");

                excelApp = Activator.CreateInstance(excelType);
                dynamic app = excelApp;
                app.Visible = false;
                app.DisplayAlerts = false;

                workbooks = app.Workbooks;
                dynamic books = workbooks;
                workbook = books.Open(path, Type.Missing, true);
                dynamic wb = workbook;

                worksheets = wb.Worksheets;
                dynamic sheets = worksheets;
                worksheet = sheets.Item[1];
                dynamic sheet = worksheet;

                usedRange = sheet.UsedRange;
                dynamic range = usedRange;
                object value2 = range.Value2;

                var values = value2 as object[,];
                if (values != null)
                {
                    int rowStart = values.GetLowerBound(0);
                    int rowEnd = values.GetUpperBound(0);
                    int colStart = values.GetLowerBound(1);
                    int colEnd = values.GetUpperBound(1);

                    for (int r = rowStart; r <= rowEnd; r++)
                    {
                        var row = new string[colEnd - colStart + 1];
                        bool hasAny = false;

                        for (int c = colStart; c <= colEnd; c++)
                        {
                            object cell = values[r, c];
                            string text = cell != null ? cell.ToString().Trim() : string.Empty;
                            row[c - colStart] = text;
                            if (!string.IsNullOrWhiteSpace(text)) hasAny = true;
                        }

                        if (hasAny)
                            table.Add(row);
                    }
                }
                else if (value2 != null)
                {
                    string text = value2.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                        table.Add(new[] { text });
                }
            }
            finally
            {
                try
                {
                    if (workbook != null)
                        ((dynamic)workbook).Close(false);
                }
                catch { }

                try
                {
                    if (excelApp != null)
                        ((dynamic)excelApp).Quit();
                }
                catch { }

                ReleaseComObject(usedRange);
                ReleaseComObject(worksheet);
                ReleaseComObject(worksheets);
                ReleaseComObject(workbook);
                ReleaseComObject(workbooks);
                ReleaseComObject(excelApp);
            }

            return ParseExceptionsFromTable(table);
        }

        private static List<TranslitException> ParseImportedXlsx(string path)
        {
            var table = new List<string[]>();

            using (var fs = File.OpenRead(path))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var shared = ReadSharedStrings(zip);
                string sheetPath = ResolveFirstWorksheetPath(zip);
                if (string.IsNullOrWhiteSpace(sheetPath))
                    return new List<TranslitException>();

                var sheetEntry = zip.GetEntry(sheetPath);
                if (sheetEntry == null)
                    return new List<TranslitException>();

                using (var stream = sheetEntry.Open())
                {
                    var doc = XDocument.Load(stream);
                    var ns = doc.Root != null ? doc.Root.Name.Namespace : XNamespace.None;
                    var rows = doc.Descendants(ns + "row");

                    foreach (var row in rows)
                    {
                        var values = new Dictionary<int, string>();
                        int maxCol = -1;

                        foreach (var cell in row.Elements(ns + "c"))
                        {
                            string cellRef = (string)cell.Attribute("r");
                            int col = GetColumnIndex(cellRef);
                            if (col < 0) continue;

                            values[col] = ReadXlsxCellValue(cell, ns, shared);
                            if (col > maxCol) maxCol = col;
                        }

                        if (maxCol < 0) continue;

                        var rowValues = new string[maxCol + 1];
                        foreach (var kv in values)
                            rowValues[kv.Key] = kv.Value;
                        table.Add(rowValues);
                    }
                }
            }

            return ParseExceptionsFromTable(table);
        }

        private static List<TranslitException> ParseExceptionsFromTable(List<string[]> table)
        {
            var result = new List<TranslitException>();
            if (table == null || table.Count == 0) return result;

            int colEnabled = 0;
            int colOriginal = 1;
            int colReplacement = 2;
            int colCategory = 3;
            int colDescription = 4;
            int startRow = 0;

            if (TryResolveTableHeader(table[0], out int hEnabled, out int hOriginal, out int hReplacement, out int hCategory, out int hDescription))
            {
                colEnabled = hEnabled;
                colOriginal = hOriginal;
                colReplacement = hReplacement;
                colCategory = hCategory;
                colDescription = hDescription;
                startRow = 1;
            }

            for (int i = startRow; i < table.Count; i++)
            {
                var row = table[i];
                if (row == null || row.Length == 0) continue;

                string original = GetTableField(row, colOriginal).Trim();
                string replacement = GetTableField(row, colReplacement).Trim();
                if (string.IsNullOrWhiteSpace(original) && string.IsNullOrWhiteSpace(replacement))
                    continue;
                if (string.IsNullOrWhiteSpace(original) || string.IsNullOrWhiteSpace(replacement))
                    continue;

                string category = GetTableField(row, colCategory);
                string description = GetTableField(row, colDescription);
                bool enabled = ParseEnabledOrDefault(GetTableField(row, colEnabled), true);

                result.Add(new TranslitException
                {
                    Original = original,
                    Replacement = replacement,
                    Category = category,
                    Description = description,
                    Enabled = enabled
                });
            }

            return result;
        }

        private static string GetTableField(string[] row, int index)
        {
            if (row == null || index < 0 || index >= row.Length) return string.Empty;
            return row[index] ?? string.Empty;
        }

        private static bool TryResolveTableHeader(
            string[] headerRow,
            out int colEnabled,
            out int colOriginal,
            out int colReplacement,
            out int colCategory,
            out int colDescription)
        {
            colEnabled = colOriginal = colReplacement = colCategory = colDescription = -1;
            if (headerRow == null || headerRow.Length == 0) return false;

            string[] headers = headerRow
                .Select(h => (h ?? string.Empty).Trim().Trim('\uFEFF').ToLowerInvariant())
                .ToArray();

            colEnabled = FindHeader(headers, "enabled", "active", "isactive", "status");
            colOriginal = FindHeader(headers, "original", "latin", "lotin");
            colReplacement = FindHeader(headers, "replacement", "cyrillic", "kiril", "kirill");
            colCategory = FindHeader(headers, "category", "kategoriya");
            colDescription = FindHeader(headers, "description", "note", "izoh", "tavsif");

            return colOriginal >= 0 && colReplacement >= 0;
        }

        private static List<string> ReadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return list;

            using (var stream = entry.Open())
            {
                var doc = XDocument.Load(stream);
                var ns = doc.Root != null ? doc.Root.Name.Namespace : XNamespace.None;
                foreach (var si in doc.Descendants(ns + "si"))
                {
                    string text = string.Concat(si.Descendants(ns + "t").Select(t => (string)t));
                    list.Add(text ?? string.Empty);
                }
            }

            return list;
        }

        private static string ResolveFirstWorksheetPath(ZipArchive zip)
        {
            var workbook = zip.GetEntry("xl/workbook.xml");
            var rels = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (workbook == null || rels == null) return null;

            XDocument workbookDoc;
            XDocument relsDoc;
            using (var stream = workbook.Open()) workbookDoc = XDocument.Load(stream);
            using (var stream = rels.Open()) relsDoc = XDocument.Load(stream);

            var wbNs = workbookDoc.Root != null ? workbookDoc.Root.Name.Namespace : XNamespace.None;
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace pkgNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            var firstSheet = workbookDoc.Descendants(wbNs + "sheet").FirstOrDefault();
            if (firstSheet == null) return null;

            string relId = (string)firstSheet.Attribute(relNs + "id");
            if (string.IsNullOrWhiteSpace(relId)) return null;

            var rel = relsDoc.Descendants(pkgNs + "Relationship")
                .FirstOrDefault(r => string.Equals((string)r.Attribute("Id"), relId, StringComparison.OrdinalIgnoreCase));
            if (rel == null) return null;

            string target = (string)rel.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target)) return null;

            target = target.Replace('\\', '/');
            if (target.StartsWith("/"))
                target = target.TrimStart('/');
            else if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                target = "xl/" + target;

            return target;
        }

        private static int GetColumnIndex(string cellRef)
        {
            if (string.IsNullOrWhiteSpace(cellRef)) return -1;

            int col = 0;
            for (int i = 0; i < cellRef.Length; i++)
            {
                char c = cellRef[i];
                if (!char.IsLetter(c)) break;
                col = (col * 26) + (char.ToUpperInvariant(c) - 'A' + 1);
            }

            return col > 0 ? col - 1 : -1;
        }

        private static string ReadXlsxCellValue(XElement cell, XNamespace ns, List<string> sharedStrings)
        {
            if (cell == null) return string.Empty;

            string type = (string)cell.Attribute("t");
            if (string.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase))
            {
                string inline = string.Concat(cell.Descendants(ns + "t").Select(t => (string)t));
                return (inline ?? string.Empty).Trim();
            }

            string raw = ((string)cell.Element(ns + "v") ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            int idx;
            if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(raw, out idx) &&
                idx >= 0 && idx < sharedStrings.Count)
            {
                return (sharedStrings[idx] ?? string.Empty).Trim();
            }

            return raw;
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null) return;
            try
            {
                if (Marshal.IsComObject(comObject))
                    Marshal.FinalReleaseComObject(comObject);
            }
            catch { }
        }

        private static string CsvEsc(string s)
        {
            if (s == null) return "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static bool ParseEnabledOrDefault(string value, bool defaultValue)
        {
            if (string.IsNullOrWhiteSpace(value)) return defaultValue;
            string v = value.Trim().ToLowerInvariant();
            if (v == "true" || v == "1" || v == "yes" || v == "y" || v == "enabled" || v == "on" || v == "faol")
                return true;
            if (v == "false" || v == "0" || v == "no" || v == "n" || v == "disabled" || v == "off" || v == "nofaol")
                return false;
            if (bool.TryParse(v, out bool b)) return b;
            return defaultValue;
        }

        private static int FindHeader(string[] headers, params string[] keys)
        {
            if (headers == null || keys == null) return -1;
            for (int i = 0; i < headers.Length; i++)
            {
                string h = headers[i];
                for (int k = 0; k < keys.Length; k++)
                {
                    if (h.Equals(keys[k], StringComparison.OrdinalIgnoreCase))
                        return i;
                }
            }
            return -1;
        }

        private static List<List<string>> ParseCsvRows(string csv)
        {
            var rows = new List<List<string>>();
            if (string.IsNullOrEmpty(csv)) return rows;

            var row = new List<string>();
            var field = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < csv.Length; i++)
            {
                char c = csv[i];

                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < csv.Length && csv[i + 1] == '"')
                        {
                            field.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        field.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else if (c == ',')
                    {
                        row.Add(field.ToString());
                        field.Clear();
                    }
                    else if (c == '\r' || c == '\n')
                    {
                        row.Add(field.ToString());
                        field.Clear();

                        if (row.Any(v => !string.IsNullOrWhiteSpace(v)))
                            rows.Add(row);
                        row = new List<string>();

                        if (c == '\r' && i + 1 < csv.Length && csv[i + 1] == '\n')
                            i++;
                    }
                    else
                    {
                        field.Append(c);
                    }
                }
            }

            row.Add(field.ToString());
            if (row.Any(v => !string.IsNullOrWhiteSpace(v)))
                rows.Add(row);

            return rows;
        }

        private static TranslitException NormalizeImportedException(TranslitException source)
        {
            if (source == null) return null;

            string original = (source.Original ?? "").Trim();
            string replacement = (source.Replacement ?? "").Trim();
            if (string.IsNullOrWhiteSpace(original) || string.IsNullOrWhiteSpace(replacement))
                return null;

            return new TranslitException
            {
                Original = original,
                Replacement = replacement,
                Category = NormalizeCategoryKey(source.Category),
                Description = string.IsNullOrWhiteSpace(source.Description) ? null : source.Description.Trim(),
                Enabled = source.Enabled,
                Note = string.IsNullOrWhiteSpace(source.Note) ? null : source.Note.Trim()
            };
        }

        private static string NormalizeCategoryKey(string category)
        {
            if (string.IsNullOrWhiteSpace(category)) return "Other";
            string value = category.Trim();

            int keyIndex = Array.FindIndex(CategoryKeys,
                k => k.Equals(value, StringComparison.OrdinalIgnoreCase));
            if (keyIndex >= 0) return CategoryKeys[keyIndex];

            int displayIndex = Array.FindIndex(Categories,
                d => d.Equals(value, StringComparison.OrdinalIgnoreCase));
            if (displayIndex >= 0) return CategoryKeys[displayIndex];

            return "Other";
        }

        private static string JsonEsc(string s)
        {
            if (s == null) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                    .Replace("\n", "\\n").Replace("\r", "\\r");
        }

        // =================================================================
        //  HELPERS
        // =================================================================

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

        // =================================================================
        //  INLINE ADD CARD — subtle surface panel with rounded corners
        // =================================================================

        private sealed class InlineAddCard : Panel
        {
            public InlineAddCard()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.UserPaint |
                         ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw, true);

                BackColor = ThemeManager.Surface;
                Height = 48;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                var rect = new Rectangle(0, 0, Width - 1, Height - 1);
                int r = 10;

                // Surface fill
                using (var path = MakeRR(rect, r))
                using (var br = new SolidBrush(ThemeManager.Surface))
                    g.FillPath(br, path);

                // Subtle border
                using (var path = MakeRR(rect, r))
                using (var pen = new Pen(ThemeManager.Border, 1f))
                    g.DrawPath(pen, path);
            }

            private static GraphicsPath MakeRR(Rectangle rect, int r)
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
}
