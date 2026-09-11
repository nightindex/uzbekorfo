using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>
    /// Floating dialog showing correction suggestions for a misspelled word.
    /// Opens near the cursor position for quick access.
    /// </summary>
    public class SuggestionsForm : ModernForm
    {
        private readonly string _errorWord;
        private readonly List<Suggestion> _suggestions;
        private readonly Action<string> _onReplace;
        private readonly Action<string> _onReplaceAll;
        private readonly Func<AddWordResult> _onAddToDict;
        private readonly bool _isGrammarError;
        private readonly string _grammarMessage;

        private ListBox _listBox;
        private Label _wordLabel;

        /// <summary>
        /// The suggestion selected by the user (null if cancelled).
        /// </summary>
        public string SelectedSuggestion { get; private set; }

        public SuggestionsForm(
            string errorWord,
            List<Suggestion> suggestions,
            Action<string> onReplace,
            Action<string> onReplaceAll = null,
            Func<AddWordResult> onAddToDict = null,
            bool isGrammarError = false,
            string grammarMessage = null)
        {
            _errorWord = errorWord;
            _suggestions = suggestions ?? new List<Suggestion>();
            _onReplace = onReplace;
            _onReplaceAll = onReplaceAll;
            _onAddToDict = onAddToDict;
            _isGrammarError = isGrammarError;
            _grammarMessage = grammarMessage;

            Title = "Вариантлар";
            Size = new Size(480, 520);
            MinimumSize = new Size(380, 400);
            ShowMinimizeButton = false;

            BuildUI();
            Shown += (s, e) => PositionNearCursor();
        }

        private void BuildUI()
        {
            // Error word display
            var accentColor = _isGrammarError
                ? Color.FromArgb(56, 142, 60)   // green for grammar
                : ThemeManager.Error;            // red for spelling

            var headerText = _isGrammarError
                ? $"Грамматика:  {_errorWord}"
                : $"Хато сўз:  {_errorWord}";

            _wordLabel = new Label
            {
                Text = headerText,
                Dock = DockStyle.Top,
                Height = _isGrammarError && !string.IsNullOrEmpty(_grammarMessage) ? 58 : 38,
                ForeColor = accentColor,
                Font = ThemeManager.FontXLBold,
                Padding = new Padding(0, 6, 0, 6)
            };

            // Show grammar message below the header word
            if (_isGrammarError && !string.IsNullOrEmpty(_grammarMessage))
            {
                _wordLabel.Text += "\n" + _grammarMessage;
            }

            // Underline simulation
            var underline = new Panel
            {
                Dock = DockStyle.Top,
                Height = 2,
                BackColor = accentColor
            };

            // Suggestions list
            _listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                BackColor = ThemeManager.Background,
                ForeColor = ThemeManager.TextPrimary,
                Font = ThemeManager.FontLG,
                ItemHeight = 42,
                DrawMode = DrawMode.OwnerDrawFixed
            };

            _listBox.DrawItem += ListBox_DrawItem;
            _listBox.DoubleClick += (s, e) => ApplySelectedReplace();
            _listBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) ApplySelectedReplace();
            };

            // Populate suggestions
            foreach (var suggestion in _suggestions)
            {
                _listBox.Items.Add(suggestion);
            }

            if (_listBox.Items.Count > 0)
                _listBox.SelectedIndex = 0;

            // Action buttons in ActionBar
            int btnH = 40;
            int btnY = (ActionBar.Height - btnH) / 2 - 2;

            var btnReplace = new ModernButton
            {
                Text = "Алмаштириш",
                Style = ModernButton.ButtonStyle.Primary,
                Size = new Size(140, btnH),
                Font = ThemeManager.FontLG,
                Location = new Point(ThemeManager.SpaceLG, btnY)
            };
            btnReplace.Click += (s, e) => ApplySelectedReplace();

            var btnReplaceAll = new ModernButton
            {
                Text = "Барчасини",
                Style = ModernButton.ButtonStyle.Secondary,
                Size = new Size(120, btnH),
                Font = ThemeManager.FontLG,
                Location = new Point(ThemeManager.SpaceLG + 150, btnY)
            };
            btnReplaceAll.Click += (s, e) => ApplyReplaceAll();

            var btnAddDict = new ModernButton
            {
                Text = "Луғатга +",
                Style = ModernButton.ButtonStyle.Ghost,
                Size = new Size(110, btnH),
                Font = ThemeManager.FontLG,
                Location = new Point(ActionBar.Width - 110 - ThemeManager.SpaceXL, btnY),
                Visible = !_isGrammarError  // hide for grammar errors
            };
            btnAddDict.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnAddDict.Click += (s, e) =>
            {
                var result = _onAddToDict?.Invoke() ?? AddWordResult.Invalid;
                if (result == AddWordResult.Invalid) return;
                Close();
            };

            ActionBar.Controls.Add(btnReplace);
            ActionBar.Controls.Add(btnReplaceAll);
            ActionBar.Controls.Add(btnAddDict);

            // Assemble
            ContentPanel.Padding = new Padding(ThemeManager.SpaceLG, ThemeManager.SpaceSM,
                                               ThemeManager.SpaceLG, 0);
            ContentPanel.Controls.Add(_listBox);
            ContentPanel.Controls.Add(underline);
            ContentPanel.Controls.Add(_wordLabel);
        }

        // =====================================================================
        //  CUSTOM LIST DRAWING
        // =====================================================================

        private void ListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var suggestion = (Suggestion)_listBox.Items[e.Index];
            bool selected = (e.State & DrawItemState.Selected) != 0;
            bool isBest = e.Index == 0;

            // Background
            var bgColor = selected ? ThemeManager.SelectedRow : ThemeManager.Background;
            using (var brush = new SolidBrush(bgColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            // Left accent for best match
            if (isBest)
            {
                using (var brush = new SolidBrush(ThemeManager.Success))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds.X, e.Bounds.Y, 3, e.Bounds.Height);
                }
            }

            // Star for best match
            int textX = e.Bounds.X + Px(12);
            if (isBest)
            {
                using (var brush = new SolidBrush(ThemeManager.Warning))
                {
                    e.Graphics.DrawString("★", UiFont(ThemeManager.FontBase), brush, textX, e.Bounds.Y + Px(8));
                }
                textX += Px(20);
            }

            // Suggestion text
            var textFont = UiFont(isBest ? ThemeManager.FontLGBold : ThemeManager.FontLG);
            using (var brush = new SolidBrush(ThemeManager.TextPrimary))
            {
                e.Graphics.DrawString(suggestion.Text, textFont, brush, textX, e.Bounds.Y + Px(7));
            }

            // Confidence percentage on right
            string confText = $"{suggestion.Confidence:P0}";
            var confSize = e.Graphics.MeasureString(confText, UiFont(ThemeManager.FontSM));
            using (var brush = new SolidBrush(ThemeManager.TextSecondary))
            {
                e.Graphics.DrawString(confText, UiFont(ThemeManager.FontSM), brush,
                    e.Bounds.Right - confSize.Width - Px(12), e.Bounds.Y + Px(10));
            }

            // Bottom border
            using (var pen = new Pen(ThemeManager.Border))
            {
                e.Graphics.DrawLine(pen, e.Bounds.X, e.Bounds.Bottom - 1,
                    e.Bounds.Right, e.Bounds.Bottom - 1);
            }
        }

        // =====================================================================
        //  ACTIONS
        // =====================================================================

        private void ApplySelectedReplace()
        {
            if (_listBox.SelectedItem is Suggestion selected)
            {
                SelectedSuggestion = selected.Text;
                _onReplace?.Invoke(selected.Text);
                Close();
            }
        }

        private void ApplyReplaceAll()
        {
            if (_listBox.SelectedItem is Suggestion selected)
            {
                SelectedSuggestion = selected.Text;
                _onReplaceAll?.Invoke(selected.Text);
                Close();
            }
        }

        // =====================================================================
        //  POSITIONING
        // =====================================================================

        private void PositionNearCursor()
        {
            try
            {
                // Try to position near the Word cursor using ActiveWindow.GetPoint
                var window = Globals.ThisAddIn.Application.ActiveWindow;
                int left, top, width, height;
                window.GetPoint(out left, out top, out width, out height,
                    Globals.ThisAddIn.Application.Selection.Range);

                var screen = Screen.FromPoint(new Point(left, top)).WorkingArea;
                int x = Math.Min(left + width + Px(10), screen.Right - Width);
                int y = Math.Min(top, screen.Bottom - Height);

                Bounds = ScreenGeometry.Fit(new Rectangle(x, y, Width, Height), screen);
                StartPosition = FormStartPosition.Manual;
            }
            catch
            {
                // Fallback to center screen
                StartPosition = FormStartPosition.CenterScreen;
            }
        }
    }
}
