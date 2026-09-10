using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Services;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>
    /// Modern export dialog with format selection, option toggles,
    /// live preview, and document statistics display.
    /// </summary>
    public class ExportForm : ModernForm
    {
        // --- Data ---
        private readonly List<ErrorEntry> _errors;
        private readonly ExportService.ExportOptions _options;

        // --- Format selector ---
        private Panel _formatPanel;
        private readonly Dictionary<ExportFormat, Panel> _formatCards = new Dictionary<ExportFormat, Panel>();
        private ExportFormat _selectedFormat = ExportFormat.Html;

        // --- Option toggles ---
        private ModernToggle _toggleContext;
        private ModernToggle _toggleSuggestions;
        private ModernToggle _toggleParagraphs;
        private ModernToggle _toggleStatistics;
        private ModernToggle _toggleOnlyUnresolved;

        // --- Preview ---
        private RichTextBox _previewBox;

        // --- Stats panel ---
        private Panel _statsPanel;

        // --- Buttons ---
        private ModernButton _btnExport;
        private ModernButton _btnCancel;

        public ExportForm(List<ErrorEntry> errors, string documentName)
        {
            _errors = errors ?? new List<ErrorEntry>();
            _options = new ExportService.ExportOptions
            {
                DocumentName = documentName ?? ""
            };

            Title = "Экспорт — Ўзбек Орфо";
            Size = new Size(780, 640);
            MinimumSize = new Size(640, 520);

            BuildUI();
            UpdatePreview();
        }

        // =====================================================================
        //  UI CONSTRUCTION
        // =====================================================================

        private void BuildUI()
        {
            // Main layout: Left (options) | Right (preview)
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 310,
                SplitterWidth = 1,
                BackColor = ThemeManager.Border,
                Panel1MinSize = 260,
                Panel2MinSize = 280,
                IsSplitterFixed = true
            };
            splitContainer.Panel1.BackColor = ThemeManager.Background;
            splitContainer.Panel2.BackColor = ThemeManager.Background;

            // ── LEFT SIDE ──
            var leftScroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(ThemeManager.SpaceLG)
            };

            int y = 0;

            // Format selection header
            var lblFormat = CreateSectionHeader("📁  Формат танланг", ref y);
            leftScroll.Controls.Add(lblFormat);
            y += 8;

            // Format cards
            _formatPanel = new Panel
            {
                Location = new Point(0, y),
                Width = 278,
                AutoSize = true
            };
            BuildFormatCards();
            leftScroll.Controls.Add(_formatPanel);
            y += _formatPanel.Height + ThemeManager.SpaceLG;

            // Options header
            var lblOptions = CreateSectionHeader("⚙️  Параметрлар", ref y);
            leftScroll.Controls.Add(lblOptions);
            y += 8;

            // Option toggles
            y = AddToggleOption(leftScroll, "Контекст кўрсатиш", true, out _toggleContext, y);
            y = AddToggleOption(leftScroll, "Тузатиш варианти", true, out _toggleSuggestions, y);
            y = AddToggleOption(leftScroll, "Параграф рақами", true, out _toggleParagraphs, y);
            y = AddToggleOption(leftScroll, "Статистика қўшиш", true, out _toggleStatistics, y);
            y = AddToggleOption(leftScroll, "Фақат тузатилмаганлар", true, out _toggleOnlyUnresolved, y);

            // Wire toggle events
            _toggleContext.Toggled += (s, e) => { _options.IncludeContext = _toggleContext.IsOn; UpdatePreview(); };
            _toggleSuggestions.Toggled += (s, e) => { _options.IncludeSuggestions = _toggleSuggestions.IsOn; UpdatePreview(); };
            _toggleParagraphs.Toggled += (s, e) => { _options.IncludeParagraphNumbers = _toggleParagraphs.IsOn; UpdatePreview(); };
            _toggleStatistics.Toggled += (s, e) => { _options.IncludeStatistics = _toggleStatistics.IsOn; UpdatePreview(); UpdateStatsPanel(); };
            _toggleOnlyUnresolved.Toggled += (s, e) => { _options.OnlyUnresolved = _toggleOnlyUnresolved.IsOn; UpdatePreview(); UpdateStatsPanel(); };

            // Statistics panel header
            y += ThemeManager.SpaceMD;
            var lblStats = CreateSectionHeader("📊  Статистика", ref y);
            leftScroll.Controls.Add(lblStats);
            y += 8;

            _statsPanel = new Panel
            {
                Location = new Point(0, y),
                Size = new Size(278, 140),
                BackColor = ThemeManager.Surface
            };
            leftScroll.Controls.Add(_statsPanel);
            UpdateStatsPanel();

            splitContainer.Panel1.Controls.Add(leftScroll);

            // ── RIGHT SIDE: Preview ──
            var rightPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(ThemeManager.SpaceLG)
            };

            var lblPreview = new Label
            {
                Text = "👁  Кўриб чиқиш",
                Font = ThemeManager.FontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                Dock = DockStyle.Top,
                Height = 28,
                Padding = new Padding(0, 0, 0, ThemeManager.SpaceSM)
            };

            _previewBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = ThemeManager.Surface,
                ForeColor = ThemeManager.TextPrimary,
                Font = new Font(ThemeManager.FontFamilyMono, 9f),
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both
            };

            rightPanel.Controls.Add(_previewBox);
            rightPanel.Controls.Add(lblPreview);
            splitContainer.Panel2.Controls.Add(rightPanel);

            ContentPanel.Padding = new Padding(0);
            ContentPanel.Controls.Add(splitContainer);

            // ── ACTION BAR ──
            _btnCancel = new ModernButton
            {
                Text = "Бекор қилиш",
                Style = ModernButton.ButtonStyle.Secondary,
                Size = new Size(130, 36),
                Anchor = AnchorStyles.Right
            };
            _btnCancel.Click += (s, e) => Close();

            _btnExport = new ModernButton
            {
                Text = "💾  Экспорт қилиш",
                Style = ModernButton.ButtonStyle.Primary,
                Size = new Size(160, 36),
                Anchor = AnchorStyles.Right
            };
            _btnExport.Click += OnExportClick;

            // Layout buttons in action bar
            var btnFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                WrapContents = false,
                Padding = new Padding(0)
            };
            btnFlow.Controls.Add(_btnExport);
            btnFlow.Controls.Add(_btnCancel);
            ActionBar.Controls.Add(btnFlow);

            // Error count label
            var filtered = _options.OnlyUnresolved
                ? _errors.Count(e => !e.IsResolved)
                : _errors.Count;
            var lblCount = new Label
            {
                Text = $"{filtered} та хато экспорт қилинади",
                Font = ThemeManager.FontBase,
                ForeColor = ThemeManager.TextSecondary,
                AutoSize = true,
                Dock = DockStyle.Left,
                Padding = new Padding(ThemeManager.SpaceSM, 10, 0, 0)
            };
            ActionBar.Controls.Add(lblCount);
        }

        // =====================================================================
        //  FORMAT CARDS
        // =====================================================================

        private void BuildFormatCards()
        {
            var formats = new[]
            {
                new { Format = ExportFormat.Text, Icon = "📝", Label = "TXT" },
                new { Format = ExportFormat.Csv, Icon = "📊", Label = "CSV" },
                new { Format = ExportFormat.Html, Icon = "🌐", Label = "HTML" },
                new { Format = ExportFormat.Json, Icon = "{ }", Label = "JSON" },
            };

            int cardWidth = 62;
            int gap = 8;
            int x = 0;

            foreach (var fmt in formats)
            {
                var card = CreateFormatCard(fmt.Format, fmt.Icon, fmt.Label, x);
                _formatPanel.Controls.Add(card);
                _formatCards[fmt.Format] = card;
                x += cardWidth + gap;
            }

            _formatPanel.Height = 72;
            HighlightSelectedFormat();
        }

        private Panel CreateFormatCard(ExportFormat format, string icon, string label, int x)
        {
            var card = new Panel
            {
                Location = new Point(x, 0),
                Size = new Size(62, 68),
                Cursor = Cursors.Hand,
                Tag = format
            };

            var lblIcon = new EmojiLabel
            {
                Text = icon,
                Font = new Font(ThemeManager.FontFamily, 18f),
                ForeColor = ThemeManager.TextPrimary,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 38,
                Cursor = Cursors.Hand
            };

            var lblName = new Label
            {
                Text = label,
                Font = ThemeManager.FontSM,
                ForeColor = ThemeManager.TextSecondary,
                TextAlign = ContentAlignment.TopCenter,
                Dock = DockStyle.Fill,
                Cursor = Cursors.Hand
            };

            card.Controls.Add(lblName);
            card.Controls.Add(lblIcon);

            // Click handlers
            EventHandler onClick = (s, e) =>
            {
                _selectedFormat = format;
                HighlightSelectedFormat();
                UpdatePreview();
            };
            card.Click += onClick;
            lblIcon.Click += onClick;
            lblName.Click += onClick;

            return card;
        }

        private void HighlightSelectedFormat()
        {
            foreach (var kv in _formatCards)
            {
                bool selected = kv.Key == _selectedFormat;
                var card = kv.Value;
                card.BackColor = selected ? ThemeManager.PrimaryLight : ThemeManager.Surface;

                // Update child label colors
                foreach (Control c in card.Controls)
                {
                    if (c is Label lbl && lbl.Font.Size < 12)
                        lbl.ForeColor = selected ? ThemeManager.Primary : ThemeManager.TextSecondary;
                }

                // Paint border
                card.Paint -= PaintCardBorder;
                card.Paint += PaintCardBorder;
                card.Invalidate();
            }
        }

        private void PaintCardBorder(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            var format = (ExportFormat)panel.Tag;
            bool selected = format == _selectedFormat;

            using (var pen = new Pen(selected ? ThemeManager.Primary : ThemeManager.Border, selected ? 2f : 1f))
            {
                var rect = new Rectangle(0, 0, panel.Width - 1, panel.Height - 1);
                using (var path = RoundedRect(rect, ThemeManager.RadiusSM))
                    e.Graphics.DrawPath(pen, path);
            }
        }

        // =====================================================================
        //  TOGGLE OPTIONS
        // =====================================================================

        private int AddToggleOption(Panel parent, string label, bool defaultOn, out ModernToggle toggle, int y)
        {
            var row = new Panel
            {
                Location = new Point(0, y),
                Size = new Size(278, 30)
            };

            var lbl = new Label
            {
                Text = label,
                Font = ThemeManager.FontBase,
                ForeColor = ThemeManager.TextPrimary,
                Location = new Point(0, 5),
                AutoSize = true
            };

            toggle = new ModernToggle
            {
                IsOn = defaultOn,
                ShowLabel = false,
                Size = new Size(44, 22),
                Location = new Point(230, 4)
            };

            row.Controls.Add(lbl);
            row.Controls.Add(toggle);
            parent.Controls.Add(row);

            return y + 34;
        }

        // =====================================================================
        //  STATISTICS PANEL
        // =====================================================================

        private void UpdateStatsPanel()
        {
            _statsPanel.Controls.Clear();
            var stats = ExportService.GetStatistics(_errors);

            int y = ThemeManager.SpaceSM;
            int x = ThemeManager.SpaceMD;

            AddStatRow("Жами хатолар:", stats.TotalErrors.ToString(), ThemeManager.Error, ref y);
            AddStatRow("Тузатилган:", stats.ResolvedErrors.ToString(), ThemeManager.Success, ref y);
            AddStatRow("Тузатилмаган:", stats.UnresolvedErrors.ToString(), ThemeManager.Warning, ref y);
            AddStatRow("Имло:", stats.SpellingErrors.ToString(), ThemeManager.TextSecondary, ref y);
            AddStatRow("Грамматика:", stats.GrammarErrors.ToString(), Color.FromArgb(56, 142, 60), ref y);
            AddStatRow("Услуб:", stats.StyleErrors.ToString(), ThemeManager.TextSecondary, ref y);
            AddStatRow("Параграфлар:", stats.AffectedParagraphs.ToString(), ThemeManager.Primary, ref y);

            _statsPanel.Height = y + ThemeManager.SpaceSM;

            // Paint rounded background
            _statsPanel.Paint -= PaintStatsBackground;
            _statsPanel.Paint += PaintStatsBackground;
            _statsPanel.Invalidate();
        }

        private void AddStatRow(string label, string value, Color valueColor, ref int y)
        {
            var lbl = new Label
            {
                Text = label,
                Font = ThemeManager.FontSM,
                ForeColor = ThemeManager.TextSecondary,
                Location = new Point(ThemeManager.SpaceMD, y),
                AutoSize = true
            };

            var val = new Label
            {
                Text = value,
                Font = ThemeManager.FontBaseBold,
                ForeColor = valueColor,
                Location = new Point(180, y),
                AutoSize = true
            };

            _statsPanel.Controls.Add(lbl);
            _statsPanel.Controls.Add(val);
            y += 20;
        }

        private void PaintStatsBackground(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(ThemeManager.Border, 1f))
            using (var path = RoundedRect(new Rectangle(0, 0, _statsPanel.Width - 1, _statsPanel.Height - 1), ThemeManager.RadiusMD))
            {
                e.Graphics.DrawPath(pen, path);
            }
        }

        // =====================================================================
        //  PREVIEW
        // =====================================================================

        private void UpdatePreview()
        {
            try
            {
                // Word format doesn't have a text preview
                if (_selectedFormat == ExportFormat.Word)
                {
                    _previewBox.Text = "(Word ҳужжат кўриб чиқиш мавжуд эмас.\nЭкспорт қилинганда янги Word ҳужжат яратилади.)";
                    return;
                }

                string preview = ExportService.GeneratePreview(
                    _errors, _selectedFormat, _options, maxLines: 40);
                _previewBox.Text = preview;
            }
            catch (Exception ex)
            {
                _previewBox.Text = $"Кўриб чиқишда хатолик: {ex.Message}";
            }
        }

        // =====================================================================
        //  EXPORT ACTION
        // =====================================================================

        private void OnExportClick(object sender, EventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                // Word format uses the existing Word report generator (ErrorList button)
                if (_selectedFormat == ExportFormat.Word)
                {
                    GenerateWordReport();
                    Close();
                    return;
                }

                string ext = ExportService.GetExtension(_selectedFormat);
                string formatLabel = ExportService.GetFormatLabel(_selectedFormat);

                using (var dialog = new SaveFileDialog())
                {
                    dialog.Title = "Хатоларни экспорт қилиш";
                    dialog.Filter = $"{formatLabel}|*{ext}|Барча файллар|*.*";
                    dialog.FileName = $"errors_{DateTime.Now:yyyyMMdd_HHmmss}{ext}";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportService.SaveToFile(dialog.FileName, _errors, _selectedFormat, _options);

                        var filtered = _options.OnlyUnresolved
                            ? _errors.Count(er => !er.IsResolved)
                            : _errors.Count;

                        ToastNotification.Success("Экспорт тугади",
                            $"{filtered} та хато {ext.TrimStart('.')} форматида сақланди.");

                        Close();
                    }
                }
            }, "Экспорт");
        }

        /// <summary>
        /// Generates a formatted Word document report (in-process, opens in new doc).
        /// </summary>
        private void GenerateWordReport()
        {
            var errors = _options.OnlyUnresolved
                ? _errors.Where(er => !er.IsResolved).ToList()
                : _errors.ToList();

            if (errors.Count == 0)
            {
                SafeExecutor.ShowInfo("Экспорт қилинадиган хатолар йўқ.");
                return;
            }

            var reportDoc = DocumentHelper.CreateNewDocument();
            var docName = _options.DocumentName;
            var stats = ExportService.GetStatistics(_errors);

            // --- TITLE ---
            var titleRange = reportDoc.Content;
            titleRange.Text = "ХАТОЛАР ҲИСОБОТИ\n";
            titleRange.Font.Name = "Times New Roman";
            titleRange.Font.Size = 18;
            titleRange.Font.Bold = 1;
            titleRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            titleRange.InsertParagraphAfter();

            // --- Metadata ---
            var metaRange = reportDoc.Paragraphs.Last.Range;
            metaRange.Font.Size = 11;
            metaRange.Font.Bold = 0;
            metaRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            metaRange.Text = $"Ҳужжат: {docName}\n" +
                             $"Сана: {DateTime.Now:dd.MM.yyyy  HH:mm}\n" +
                             $"Жами хатолар: {errors.Count} та\n";

            // --- Statistics block ---
            if (_options.IncludeStatistics)
            {
                metaRange.InsertParagraphAfter();
                var statsRange = reportDoc.Paragraphs.Last.Range;
                statsRange.Font.Size = 10;
                statsRange.Font.Italic = 1;
                statsRange.Text = $"Имло: {stats.SpellingErrors} | Грамматика: {stats.GrammarErrors} | Услуб: {stats.StyleErrors} | " +
                                  $"Тахминий: {stats.TypoErrors} | Параграфлар: {stats.AffectedParagraphs}\n";
            }

            metaRange.InsertParagraphAfter();

            // --- TABLE ---
            var tableRange = reportDoc.Paragraphs.Last.Range;

            // Determine column count dynamically
            var columnDefs = new List<Tuple<string, float>> { Tuple.Create("№", 30f), Tuple.Create("Сўз", 90f), Tuple.Create("Тури", 55f) };
            if (_options.IncludeSuggestions)
                columnDefs.Add(Tuple.Create("Тузатиш варианти", 90f));
            if (_options.IncludeParagraphNumbers)
                columnDefs.Add(Tuple.Create("Параграф", 55f));
            if (_options.IncludeContext)
                columnDefs.Add(Tuple.Create("Контекст", 220f));

            int rows = errors.Count + 1;
            int cols = columnDefs.Count;

            var table = reportDoc.Tables.Add(tableRange, rows, cols);
            table.Borders.Enable = 1;
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 11;

            // Set column widths
            for (int c = 0; c < cols; c++)
                table.Columns[c + 1].Width = columnDefs[c].Item2;

            // Header row
            var headerRow = table.Rows[1];
            headerRow.Range.Font.Bold = 1;
            headerRow.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorGray15;
            for (int c = 0; c < cols; c++)
                table.Cell(1, c + 1).Range.Text = columnDefs[c].Item1;

            // Data rows
            for (int i = 0; i < errors.Count; i++)
            {
                var err = errors[i];
                int row = i + 2;
                int col = 1;

                table.Cell(row, col).Range.Text = (i + 1).ToString();
                col++;

                table.Cell(row, col).Range.Text = err.Word ?? "";
                table.Cell(row, col).Range.Font.Color = err.IsGrammarError
                    ? (Microsoft.Office.Interop.Word.WdColor)3968568  // RGB(56,142,60) green
                    : Microsoft.Office.Interop.Word.WdColor.wdColorRed;
                col++;

                // Type column
                table.Cell(row, col).Range.Text = err.IsGrammarError ? "Грамматика" : "Имло";
                if (err.IsGrammarError)
                    table.Cell(row, col).Range.Font.Color = (Microsoft.Office.Interop.Word.WdColor)3968568;
                col++;

                if (_options.IncludeSuggestions)
                {
                    table.Cell(row, col).Range.Text = err.BestSuggestion ?? "—";
                    col++;
                }
                if (_options.IncludeParagraphNumbers)
                {
                    table.Cell(row, col).Range.Text = $"§{err.ParagraphIndex}";
                    col++;
                }
                if (_options.IncludeContext)
                {
                    table.Cell(row, col).Range.Text = err.Context ?? "";
                    table.Cell(row, col).Range.Font.Size = 10;
                }
            }

            // --- FOOTER ---
            var footerRange = reportDoc.Paragraphs.Last.Range;
            footerRange.InsertParagraphAfter();
            footerRange = reportDoc.Paragraphs.Last.Range;
            footerRange.Font.Size = 9;
            footerRange.Font.Italic = 1;
            footerRange.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorGray50;
            footerRange.Text = $"\nЎзбек Орфо — автоматик хатолар ҳисоботи | {DateTime.Now:dd.MM.yyyy HH:mm:ss}";

            ToastNotification.Success("Ҳисобот тайёр", $"{errors.Count} та хато рўйхати яратилди.");
        }

        // =====================================================================
        //  SECTION HEADER HELPER
        // =====================================================================

        private Label CreateSectionHeader(string text, ref int y)
        {
            var lbl = new Label
            {
                Text = text,
                Font = ThemeManager.FontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                Location = new Point(0, y),
                AutoSize = true
            };
            y += 28;
            return lbl;
        }

        // =====================================================================
        //  ROUNDED RECT HELPER
        // =====================================================================

        private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int d = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
