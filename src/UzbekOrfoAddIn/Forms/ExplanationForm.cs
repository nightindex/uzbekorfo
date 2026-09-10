using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>
    /// Modern popup form that displays the explanation/definition for a selected word.
    /// Shows: word, definition, spelling rule, grammar note, and example sentences.
    /// Positioned near the Word cursor for context-aware UX.
    /// </summary>
    public class ExplanationForm : ModernForm
    {
        private const int CloseButtonWidth = 116;
        private const int CloseButtonHeight = 40;
        private static readonly Font SectionBodyFont = new Font(ThemeManager.FontFamily, 11.25f, FontStyle.Regular);
        private static readonly Font SectionHeaderFont = new Font(ThemeManager.FontFamily, 11.25f, FontStyle.Bold);
        private static readonly Font SectionHeaderLargeFont = new Font(ThemeManager.FontFamily, 14.5f, FontStyle.Bold);
        private static readonly Font ExampleFont = new Font(ThemeManager.FontFamily, 11.25f, FontStyle.Italic);
        private static readonly Font EmptyHintFont = new Font(ThemeManager.FontFamily, 12.0f, FontStyle.Regular);

        private readonly ExplanationEntry _entry;
        private readonly string _queryWord;
        private ModernScrollBar _contentScrollBar;

        public ExplanationForm(string queryWord, ExplanationEntry entry)
        {
            _queryWord = queryWord ?? "";
            _entry = entry;

            Title = "Изоҳ — " + (_entry?.Word ?? _queryWord);
            Size = new Size(700, 720);
            MinimumSize = new Size(560, 500);
            ShowMinimizeButton = false;

            BuildUI();
            PositionNearCursor();
        }

        private void BuildUI()
        {
            if (_entry == null)
            {
                BuildNoDataUI();
                return;
            }

            var container = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(0)
            };

            int y = 0;

            // === Word Header ===
            var wordPanel = CreateSection(
                _entry.Word,
                null,
                ThemeManager.Primary,
                SectionHeaderLargeFont,
                ref y);
            container.Controls.Add(wordPanel);

            // === Definition ===
            if (!string.IsNullOrEmpty(_entry.Definition))
            {
                var defPanel = CreateSection(
                    "Маъноси",
                    _entry.Definition,
                    ThemeManager.Primary,
                    SectionHeaderFont,
                    ref y);
                container.Controls.Add(defPanel);
            }

            // === Spelling Rule ===
            if (!string.IsNullOrEmpty(_entry.SpellingRule))
            {
                var rulePanel = CreateSection(
                    "Имло қоидаси",
                    _entry.SpellingRule,
                    ThemeManager.Warning,
                    SectionHeaderFont,
                    ref y);
                container.Controls.Add(rulePanel);
            }

            // === Grammar Note ===
            if (!string.IsNullOrEmpty(_entry.GrammarNote))
            {
                var gramPanel = CreateSection(
                    "Грамматика",
                    _entry.GrammarNote,
                    ThemeManager.Accent,
                    SectionHeaderFont,
                    ref y);
                container.Controls.Add(gramPanel);
            }

            // === Examples ===
            if (_entry.Examples != null && _entry.Examples.Length > 0)
            {
                var exPanel = CreateExamplesSection(ref y);
                container.Controls.Add(exPanel);
            }

            // === Action Bar ===
            int btnY = (ActionBar.Height - CloseButtonHeight) / 2;

            var btnClose = new ModernButton
            {
                Text = "Ёпиш",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = ThemeManager.FontLGBold,
                Size = new Size(CloseButtonWidth, CloseButtonHeight),
                Location = new Point(ActionBar.Width - CloseButtonWidth - ThemeManager.SpaceXL, btnY),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnClose.Click += (s, e) => Close();
            ActionBar.Controls.Add(btnClose);

            ContentPanel.Controls.Add(container);
            AttachContentScrollBar(container);
        }

        /// <summary>
        /// Builds "no data found" UI when there's no explanation for the word.
        /// </summary>
        private void BuildNoDataUI()
        {
            Size = new Size(620, 400);

            var container = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(ThemeManager.SpaceXL, ThemeManager.SpaceLG,
                                      ThemeManager.SpaceXL, ThemeManager.SpaceLG)
            };

            // Icon
            var lblIcon = new Label
            {
                Text = "ⓘ",
                Font = new Font("Segoe UI Symbol", 46f, FontStyle.Regular),
                ForeColor = ThemeManager.Primary,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 100
            };

            // Message
            var lblMsg = new Label
            {
                Text = $"«{_queryWord}» сўзи учун изоҳ топилмади.",
                Font = SectionHeaderLargeFont,
                ForeColor = ThemeManager.TextSecondary,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 64
            };

            // Hint
            var lblHint = new Label
            {
                Text = "Бу сўз изоҳлар базасига киритилмаган.\nКейинчалик қўшилиши мумкин.",
                Font = EmptyHintFont,
                ForeColor = ThemeManager.TextSecondary,
                TextAlign = ContentAlignment.TopCenter,
                Dock = DockStyle.Top,
                Height = 84
            };

            container.Controls.Add(lblHint);
            container.Controls.Add(lblMsg);
            container.Controls.Add(lblIcon);

            // Close button
            int btnY = (ActionBar.Height - CloseButtonHeight) / 2;
            var btnClose = new ModernButton
            {
                Text = "Ёпиш",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = ThemeManager.FontLGBold,
                Size = new Size(CloseButtonWidth, CloseButtonHeight),
                Location = new Point(ActionBar.Width - CloseButtonWidth - ThemeManager.SpaceXL, btnY),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnClose.Click += (s, e) => Close();
            ActionBar.Controls.Add(btnClose);

            ContentPanel.Controls.Add(container);
        }

        private void AttachContentScrollBar(Panel container)
        {
            DisposeContentScrollBar();
            if (container == null) return;

            _contentScrollBar = ModernScrollBar.AttachTo(container);
            if (_contentScrollBar != null)
            {
                Action show = () =>
                {
                    try { _contentScrollBar.ShowScrollBar(); } catch { }
                };

                if (IsHandleCreated) show();
                else Shown += (s, e) => show();
            }
        }

        private void DisposeContentScrollBar()
        {
            if (_contentScrollBar != null)
            {
                try { _contentScrollBar.Dispose(); } catch { }
                _contentScrollBar = null;
            }
        }

        // =====================================================================
        //  SECTION BUILDERS
        // =====================================================================

        /// <summary>
        /// Creates a labeled section with a colored accent bar, header, and body text.
        /// </summary>
        private Panel CreateSection(string header, string body, Color accentColor,
            Font headerFont, ref int yPos)
        {
            bool isWordHeader = body == null; // The first section (word itself) has no body
            int contentLeft = isWordHeader ? ThemeManager.SpaceLG : ThemeManager.SpaceLG + 14;

            var panel = new Panel
            {
                Location = new Point(0, yPos),
                Width = ContentPanel.Width - 50,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                Padding = new Padding(ThemeManager.SpaceLG, ThemeManager.SpaceSM,
                                      ThemeManager.SpaceMD, ThemeManager.SpaceSM)
            };

            Panel accentBar = null;
            if (!isWordHeader)
            {
                // Keep section design minimal: just the accent line and text.
                accentBar = new Panel
                {
                    Dock = DockStyle.Left,
                    Width = 4,
                    BackColor = accentColor
                };
            }

            // Header label
            var lblHeader = new Label
            {
                Text = header,
                Font = headerFont,
                ForeColor = isWordHeader ? ThemeManager.Primary : ThemeManager.TextPrimary,
                AutoSize = true,
                MaximumSize = new Size(panel.Width - contentLeft - ThemeManager.SpaceMD, 0),
                Location = new Point(contentLeft, ThemeManager.SpaceSM)
            };

            int height = lblHeader.PreferredHeight + ThemeManager.SpaceSM * 2;

            if (!string.IsNullOrEmpty(body))
            {
                var lblBody = new Label
                {
                    Text = body,
                    Font = SectionBodyFont,
                    ForeColor = ThemeManager.TextSecondary,
                    AutoSize = true,
                    MaximumSize = new Size(panel.Width - contentLeft - ThemeManager.SpaceMD, 0),
                    Location = new Point(contentLeft,
                        lblHeader.PreferredHeight + ThemeManager.SpaceMD)
                };

                panel.Controls.Add(lblBody);
                height = lblBody.Bottom + ThemeManager.SpaceSM;
            }

            panel.Height = height;
            panel.Controls.Add(lblHeader);
            if (accentBar != null)
                panel.Controls.Add(accentBar);

            yPos += height + ThemeManager.SpaceSM;
            return panel;
        }

        /// <summary>
        /// Creates the examples section with numbered example sentences.
        /// </summary>
        private Panel CreateExamplesSection(ref int yPos)
        {
            var panel = new Panel
            {
                Location = new Point(0, yPos),
                Width = ContentPanel.Width - 50,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                Padding = new Padding(ThemeManager.SpaceLG, ThemeManager.SpaceSM,
                                      ThemeManager.SpaceMD, ThemeManager.SpaceSM)
            };

            // Accent bar
            var accentBar = new Panel
            {
                Dock = DockStyle.Left,
                Width = 4,
                BackColor = ThemeManager.Success
            };

            // Header
            var lblHeader = new Label
            {
                Text = "Мисоллар",
                Font = SectionHeaderFont,
                ForeColor = ThemeManager.TextPrimary,
                AutoSize = true,
                Location = new Point(ThemeManager.SpaceLG + 14, ThemeManager.SpaceSM)
            };

            int currentY = lblHeader.PreferredHeight + ThemeManager.SpaceMD;

            for (int i = 0; i < _entry.Examples.Length; i++)
            {
                var lblNum = new Label
                {
                    Text = $"{i + 1}.",
                    Font = SectionHeaderFont,
                    ForeColor = ThemeManager.Success,
                    AutoSize = true,
                    Location = new Point(ThemeManager.SpaceLG + 14, currentY)
                };

                var lblExample = new Label
                {
                    Text = _entry.Examples[i],
                    Font = ExampleFont,
                    ForeColor = ThemeManager.TextSecondary,
                    AutoSize = true,
                    MaximumSize = new Size(panel.Width - 60, 0),
                    Location = new Point(ThemeManager.SpaceLG + 40, currentY)
                };

                panel.Controls.Add(lblNum);
                panel.Controls.Add(lblExample);

                currentY = lblExample.Bottom + ThemeManager.SpaceSM;
            }

            panel.Height = currentY + ThemeManager.SpaceSM;
            panel.Controls.Add(lblHeader);
            panel.Controls.Add(accentBar);

            yPos += panel.Height + ThemeManager.SpaceSM;
            return panel;
        }

        // =====================================================================
        //  POSITIONING
        // =====================================================================

        /// <summary>
        /// Positions the form near the Word cursor for contextual display.
        /// Falls back to center screen if cursor position can't be determined.
        /// </summary>
        private void PositionNearCursor()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;

                if (sel != null && sel.Range != null)
                {
                    // Use ActiveWindow.GetPoint which works on the selection range
                    int left = 0, top = 0, width = 0, height = 0;
                    app.ActiveWindow.GetPoint(out left, out top, out width, out height, sel.Range);

                    var screen = Screen.PrimaryScreen.WorkingArea;
                    int x = left + width + 10;
                    int y = top;

                    // Keep on screen
                    if (x + Width > screen.Right) x = screen.Right - Width - 10;
                    if (y + Height > screen.Bottom) y = top - Height - 10;
                    if (x < screen.Left) x = screen.Left + 10;
                    if (y < screen.Top) y = screen.Top + 10;

                    StartPosition = FormStartPosition.Manual;
                    Location = new Point(x, y);
                    return;
                }
            }
            catch { }

            // Fallback: center screen
            StartPosition = FormStartPosition.CenterScreen;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            DisposeContentScrollBar();
            base.OnFormClosed(e);
        }
    }
}
