using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Services;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>
    /// Modern dialog for viewing, editing, and managing the user's personal dictionary.
    /// Features alphabetical sections, inline delete, search/filter, bulk operations,
    /// and a description panel showing word definitions and explanations.
    /// </summary>
    public class EditDictionaryForm : ModernForm
    {
        private readonly Action<string> _onAddWord;
        private readonly Action<string> _onRemoveWord;
        private readonly Func<List<string>> _getWords;
        private readonly Func<string, bool> _isMainWord;
        private readonly Action<DictionaryService.EditorWordCache> _onCacheBuilt;
        private readonly IExplanationProvider _explanationProvider;

        private ListView _wordList;
        private ModernTextBox _searchBox;
        private ModernTextBox _newWordBox;
        private Label _countLabel;
        private List<string> _allWords = new List<string>();
        private List<string> _cyrWords  = new List<string>();  // pre-split cache
        private List<string> _latWords  = new List<string>();  // pre-split cache
        private Dictionary<string, string> _latToCyrMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private List<string> _filteredWords = new List<string>();
        private ImageList _wordListRowHeightImageList;
        private Timer _searchDebounce;
        private ModernScrollBar _wordListScrollBar;
        private ModernScrollBar _descriptionScrollBar;
        private const int DefinitionScrollBarWidth = 8;
        private ModernButton _btnExportAction;
        private ModernButton _btnImportAction;
        private bool _isBusy;

        // Script-filter pills
        private Panel _pillBar;
        private int _activePillIndex; // 0 = All, 1 = Latin, 2 = Cyrillic
        private readonly List<Label> _pills = new List<Label>();
        private Panel _busyOverlay;
        private Panel _busyCard;
        private Label _busyTitleLabel;
        private Label _busyDetailLabel;
        private ProgressBar _busyProgress;

        // Description panel controls
        private Panel _descriptionPanel;
        private Label _descWordLabel;
        private EmojiLabel _descDefinitionLabel;
        private Label _descDefinitionText;
        private EmojiLabel _descSpellingLabel;
        private Label _descSpellingText;
        private EmojiLabel _descGrammarLabel;
        private Label _descGrammarText;
        private EmojiLabel _descExamplesLabel;
        private Label _descExamplesText;
        private Label _descNoDataLabel;
        private FlowLayoutPanel _descViewFlow;
        private TableLayoutPanel _descDefinitionSection;
        private TableLayoutPanel _descSpellingSection;
        private TableLayoutPanel _descGrammarSection;
        private TableLayoutPanel _descExamplesSection;

        // Edit mode controls
        private bool _isEditMode;
        private Panel _editPanel;
        private TextBox _editDefinition;
        private TextBox _editSpellingRule;
        private TextBox _editGrammarNote;
        private TextBox _editExamples;
        private ModernButton _btnEditSave;
        private ModernButton _btnCancelEdit;

        // Larger fonts for this form
        private static readonly Font _formFontLG = new Font("Segoe UI", 11.25f, FontStyle.Regular);
        private static readonly Font _fPill  = new Font("Segoe UI Semibold", 10.5f, FontStyle.Bold);
        private static readonly Font _fPillR = new Font("Segoe UI", 10.5f, FontStyle.Regular);
        private static readonly Font _formFontLGBold = new Font("Segoe UI", 11.25f, FontStyle.Bold);
        private static readonly Font _formFontXL = new Font("Segoe UI", 14f, FontStyle.Bold);
        private static readonly Font _formFont2XL = new Font("Segoe UI", 16f, FontStyle.Bold);

        public EditDictionaryForm(
            Func<List<string>> getWords,
            Action<string> onAddWord,
            Action<string> onRemoveWord,
            Func<string, bool> isMainWord = null,
            Action<DictionaryService.EditorWordCache> onCacheBuilt = null,
            IExplanationProvider explanationProvider = null,
            DictionaryService.EditorWordCache cachedData = null)
        {
            _getWords = getWords;
            _onAddWord = onAddWord;
            _onRemoveWord = onRemoveWord;
            _isMainWord = isMainWord;
            _onCacheBuilt = onCacheBuilt;
            _explanationProvider = explanationProvider;

            Title = "\u041b\u0443\u0493\u0430\u0442";
            Size = new Size(1220, 860);
            MinimumSize = new Size(980, 720);

            SuspendLayout();
            BuildUI();
            ResumeLayout(false);
            PerformLayout();

            if (cachedData != null)
            {
                // Cached data from a previous open — populate instantly.
                _allWords = cachedData.AllWords;
                _cyrWords = cachedData.CyrillicWords;
                _latWords = cachedData.LatinWords;
                _latToCyrMap = cachedData.LatinToCyrillicMap;
                FilterWords();
            }
            else
            {
                // First open: load on background System.Threading.Thread after Shown fires so the
                // dialog appears while data is being fetched.
                Shown += _shownHandler;
            }
        }

        private void _shownHandler(object s, EventArgs ev)
        {
            QueueInitialWordLoad();
        }

        /// <summary>
        /// Re-populates the form with fresh or cached data without rebuilding
        /// the entire UI.  Called before re-showing a cached form instance.
        /// </summary>
        public void Reinitialize(DictionaryService.EditorWordCache cachedData)
        {
            // Reset transient state
            _searchBox.Text = "";
            _activePillIndex = 0;
            UpdatePillStyles();

            if (_isEditMode) ExitEditMode();
            UpdateDescription(null); // clear description panel

            if (cachedData != null)
            {
                _allWords = cachedData.AllWords;
                _cyrWords = cachedData.CyrillicWords;
                _latWords = cachedData.LatinWords;
                _latToCyrMap = cachedData.LatinToCyrillicMap;
                // Unhook the background loader — data is already available
                Shown -= _shownHandler;
                FilterWords();
            }
            else
            {
                // Data was invalidated — reload on Shown
                _allWords = new List<string>();
                _cyrWords = new List<string>();
                _latWords = new List<string>();
                _latToCyrMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                _filteredWords = _allWords;
                _countLabel.Text = "0 \u0441\u045e\u0437";
                _wordList.VirtualListSize = 0;
                Shown -= _shownHandler;
                Shown += _shownHandler;
            }
        }

        private void BuildUI()
        {
            // === Top: Search + Add ===
            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 148,
                Padding = new Padding(0, 0, 0, ThemeManager.SpaceSM)
            };

            _searchBox = new ModernTextBox
            {
                Placeholder = "Qidirish / \u049A\u0438\u0434\u0438\u0440\u0438\u0448...",
                Font = _formFontLG,
                Location = new Point(0, 0),
                Size = new Size(340, 40)
            };
            // Debounced search: wait 200ms after last keystroke before filtering 80K words
            _searchDebounce = new Timer { Interval = 200 };
            _searchDebounce.Tick += (s, e) => { _searchDebounce.Stop(); FilterWords(); };
            _searchBox.TextChanged += (s, e) => { _searchDebounce.Stop(); _searchDebounce.Start(); };

            _countLabel = new Label
            {
                Text = "0 \u0441\u045e\u0437",
                AutoSize = true,
                ForeColor = ThemeManager.TextSecondary,
                Font = _formFontLG,
                Location = new Point(350, 10)
            };

            _newWordBox = new ModernTextBox
            {
                Placeholder = "\u042f\u043d\u0433\u0438 \u0441\u045e\u0437 \u049b\u045e\u0448\u0438\u0448...",
                Font = _formFontLG,
                Location = new Point(0, 50),
                Size = new Size(340, 40)
            };

            var btnAdd = new ModernButton
            {
                Text = "+ \u049a\u045e\u0448\u0438\u0448",
                Style = ModernButton.ButtonStyle.Primary,
                Font = _formFontLG,
                Size = new Size(120, 40),
                Location = new Point(350, 50)
            };
            btnAdd.Click += BtnAdd_Click;

            topPanel.Controls.Add(_searchBox);
            topPanel.Controls.Add(_countLabel);
            topPanel.Controls.Add(_newWordBox);
            topPanel.Controls.Add(btnAdd);

            // === Script-filter pill bar ===
            _pillBar = new Panel
            {
                Location = new Point(0, 100),
                Size = new Size(500, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.Transparent
            };
            BuildPills();
            topPanel.Controls.Add(_pillBar);

            // === Word list (left side) ===
            _wordList = new ListView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                BackColor = ThemeManager.Background,
                ForeColor = ThemeManager.TextPrimary,
                Font = _formFontLG,
                View = View.Details,
                HeaderStyle = ColumnHeaderStyle.None,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false,
                VirtualMode = true,
                VirtualListSize = 0
            };
            _wordList.Columns.Add("", 100);
            _wordList.RetrieveVirtualItem += WordList_RetrieveVirtualItem;
            _wordList.KeyDown += WordList_KeyDown;
            _wordList.SelectedIndexChanged += WordList_SelectedIndexChanged;
            _wordList.Resize += (s, e) => ResizeWordListColumn();

            // Force row height close to the previous ListBox item height (38px).
            _wordListRowHeightImageList = new ImageList
            {
                ImageSize = new Size(1, 38),
                ColorDepth = ColorDepth.Depth8Bit
            };
            _wordListRowHeightImageList.Images.Add(new Bitmap(1, 38));
            _wordList.SmallImageList = _wordListRowHeightImageList;
            ResizeWordListColumn();

            // === Left panel (word list container) ===
            var leftPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0)
            };
            leftPanel.Controls.Add(_wordList);
            leftPanel.Controls.Add(topPanel);
            _wordListScrollBar = ModernScrollBar.AttachTo(_wordList);

            // === Description panel (right side) ===
            BuildDescriptionPanel();

            // === Action bar (taller for bigger buttons) ===
            ActionBar.Height = 68;
            int btnH = 40;
            int btnY = (ActionBar.Height - btnH) / 2 - 2;

            var btnDelete = new ModernButton
            {
                Text = "\u040e\u0447\u0438\u0440\u0438\u0448",
                Style = ModernButton.ButtonStyle.Danger,
                Font = _formFontLG,
                Size = new Size(130, btnH),
                Location = new Point(ThemeManager.SpaceLG, btnY)
            };
            btnDelete.Click += BtnDelete_Click;

            _btnExportAction = new ModernButton
            {
                Text = "\u042d\u043a\u0441\u043f\u043e\u0440\u0442",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = _formFontLG,
                Size = new Size(130, btnH),
                Location = new Point(ThemeManager.SpaceLG + 140, btnY)
            };
            _btnExportAction.Click += BtnExport_Click;

            _btnImportAction = new ModernButton
            {
                Text = "\u0418\u043c\u043f\u043e\u0440\u0442",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = _formFontLG,
                Size = new Size(130, btnH),
                Location = new Point(ThemeManager.SpaceLG + 280, btnY)
            };
            _btnImportAction.Click += BtnImport_Click;

            var btnClose = new ModernButton
            {
                Text = "\u0401\u043f\u0438\u0448",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = _formFontLG,
                Size = new Size(100, btnH),
                Location = new Point(ActionBar.Width - 100 - ThemeManager.SpaceXL, btnY)
            };
            btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnClose.Click += (s, e) => Close();

            ActionBar.Controls.Add(btnDelete);
            ActionBar.Controls.Add(_btnExportAction);
            ActionBar.Controls.Add(_btnImportAction);
            ActionBar.Controls.Add(btnClose);

            // === Splitter between list and description ===
            var splitter = new Splitter
            {
                Dock = DockStyle.Right,
                Width = 3,
                BackColor = ThemeManager.Border,
                MinSize = 220
            };

            // === Assemble ===
            ContentPanel.Padding = new Padding(ThemeManager.SpaceLG, ThemeManager.SpaceSM,
                                               ThemeManager.SpaceLG, 0);
            ContentPanel.Controls.Add(leftPanel);
            ContentPanel.Controls.Add(splitter);
            ContentPanel.Controls.Add(_descriptionPanel);
            BuildBusyOverlay();
        }

        private void BuildBusyOverlay()
        {
            _busyOverlay = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeManager.Background,
                Visible = false
            };

            _busyCard = new Panel
            {
                Size = new Size(390, 150),
                BackColor = ThemeManager.Surface,
                BorderStyle = BorderStyle.FixedSingle
            };

            _busyTitleLabel = new Label
            {
                Text = "\u0418\u0448\u043b\u0430\u043d\u043c\u043e\u049b\u0434\u0430...",
                Font = _formFontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                Location = new Point(18, 18),
                Size = new Size(354, 28)
            };

            _busyDetailLabel = new Label
            {
                Text = "\u0418\u043b\u0442\u0438\u043c\u043e\u0441, \u043a\u0443\u0442\u0438\u043d\u0433...",
                Font = _formFontLG,
                ForeColor = ThemeManager.TextSecondary,
                Location = new Point(18, 50),
                Size = new Size(354, 24)
            };

            _busyProgress = new ProgressBar
            {
                Location = new Point(18, 92),
                Size = new Size(354, 18),
                Style = ProgressBarStyle.Blocks,
                MarqueeAnimationSpeed = 30
            };

            _busyCard.Controls.Add(_busyTitleLabel);
            _busyCard.Controls.Add(_busyDetailLabel);
            _busyCard.Controls.Add(_busyProgress);

            _busyOverlay.Controls.Add(_busyCard);
            _busyOverlay.Resize += (s, e) => CenterBusyCard();
            ContentPanel.Controls.Add(_busyOverlay);
            _busyOverlay.BringToFront();
            CenterBusyCard();
        }

        private void CenterBusyCard()
        {
            if (_busyOverlay == null || _busyCard == null) return;

            int x = Math.Max(0, (_busyOverlay.ClientSize.Width - _busyCard.Width) / 2);
            int y = Math.Max(0, (_busyOverlay.ClientSize.Height - _busyCard.Height) / 2);
            _busyCard.Location = new Point(x, y);
        }

        // =====================================================================
        //  DESCRIPTION PANEL
        // =====================================================================

        private void BuildDescriptionPanel()
        {
            _descriptionPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 420,
                BackColor = ThemeManager.Surface,
                Padding = new Padding(ThemeManager.SpaceXL, ThemeManager.SpaceLG,
                                      ThemeManager.SpaceXL, ThemeManager.SpaceLG),
                AutoScroll = false
            };

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 1f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, ThemeManager.SpaceSM));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            // Header: selected word
            var headerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };

            _descWordLabel = new Label
            {
                Text = "\u0421\u045e\u0437\u043d\u0438 \u0442\u0430\u043d\u043b\u0430\u043d\u0433...",
                Font = _formFont2XL,
                ForeColor = ThemeManager.TextSecondary,
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            headerPanel.Controls.Add(_descWordLabel);

            // Edit/Add button РІР‚вЂќ full-width under header for easy visibility
            _btnEditSave = new ModernButton
            {
                Text = "\u0422\u0430\u04b3\u0440\u0438\u0440",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = _formFontLG,
                Size = new Size(360, 36),
                Location = new Point(0, 3),
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                CornerRadius = 10,
                Margin = new Padding(0),
                Visible = false
            };
            _btnEditSave.Click += BtnEditSave_Click;

            var editButtonHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            editButtonHost.Controls.Add(_btnEditSave);
            editButtonHost.Resize += (s, e) =>
            {
                int w = Math.Max(Px(220), editButtonHost.ClientSize.Width);
                _btnEditSave.Width = w;
                _btnEditSave.Left = 0;
                _btnEditSave.Top = Math.Max(0, (editButtonHost.ClientSize.Height - _btnEditSave.Height) / 2);
            };

            // Separator
            var separator = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ThemeManager.Border
            };

            // Spacer after separator
            var spacerAfterSep = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };

            // "No data" / "Add description" label
            _descNoDataLabel = new Label
            {
                Text = "\u0411\u0443 \u0441\u045e\u0437 \u0443\u0447\u0443\u043d \u0438\u0437\u043e\u04b3 \u043c\u0430\u0432\u0436\u0443\u0434 \u044d\u043c\u0430\u0441.",
                Font = _formFontLG,
                ForeColor = ThemeManager.TextSecondary,
                AutoSize = true,
                MaximumSize = new Size(380, 0),
                Padding = new Padding(0, ThemeManager.SpaceMD, 0, 0),
                Visible = false
            };

            // === VIEW MODE labels (using EmojiLabel for color emoji support) ===
            _descDefinitionLabel = CreateSectionLabel("\uD83D\uDCD6  \u041c\u0430\u044a\u043d\u043e\u0441\u0438");
            _descDefinitionText = CreateSectionText();

            _descSpellingLabel = CreateSectionLabel("\u270F\uFE0F  \u0418\u043c\u043b\u043e \u049b\u043e\u0438\u0434\u0430\u0441\u0438");
            _descSpellingText = CreateSectionText();

            _descGrammarLabel = CreateSectionLabel("\uD83D\uDCDD  \u0413\u0440\u0430\u043c\u043c\u0430\u0442\u0438\u043a\u0430");
            _descGrammarText = CreateSectionText();

            _descExamplesLabel = CreateSectionLabel("\uD83D\uDCA1  \u041c\u0438\u0441\u043e\u043b\u043b\u0430\u0440");
            _descExamplesText = CreateSectionText();

            _descDefinitionSection = CreateSectionBlock(_descDefinitionLabel, _descDefinitionText);
            _descSpellingSection = CreateSectionBlock(_descSpellingLabel, _descSpellingText);
            _descGrammarSection = CreateSectionBlock(_descGrammarLabel, _descGrammarText);
            _descExamplesSection = CreateSectionBlock(_descExamplesLabel, _descExamplesText);

            _descViewFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Margin = new Padding(0),
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };
            _descViewFlow.Controls.Add(_descNoDataLabel);
            _descViewFlow.Controls.Add(_descDefinitionSection);
            _descViewFlow.Controls.Add(_descSpellingSection);
            _descViewFlow.Controls.Add(_descGrammarSection);
            _descViewFlow.Controls.Add(_descExamplesSection);
            _descViewFlow.SizeChanged += (s, e) => UpdateDescriptionViewWidths();

            // === EDIT MODE panel ===
            BuildEditPanel();
            _editPanel.Dock = DockStyle.Fill;

            var contentHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            contentHost.Controls.Add(_descViewFlow);
            contentHost.Controls.Add(_editPanel);
            _editPanel.Visible = false;

            root.Controls.Add(headerPanel, 0, 0);
            root.Controls.Add(separator, 0, 1);
            root.Controls.Add(editButtonHost, 0, 2);
            root.Controls.Add(spacerAfterSep, 0, 3);
            root.Controls.Add(contentHost, 0, 4);

            _descriptionPanel.Controls.Add(root);
            UpdateDescriptionViewWidths();
            AttachDescriptionScrollBar();
        }

        private void AttachDescriptionScrollBar()
        {
            if (_descViewFlow == null) return;

            _descViewFlow.Scroll -= DescViewFlow_Scroll;
            _descViewFlow.MouseWheel -= DescViewFlow_MouseWheel;

            if (_descriptionScrollBar != null)
            {
                try { _descriptionScrollBar.Dispose(); } catch { }
                _descriptionScrollBar = null;
            }

            _descriptionScrollBar = ModernScrollBar.AttachTo(_descViewFlow, DefinitionScrollBarWidth);
            if (_descriptionScrollBar != null)
            {
                _descViewFlow.Scroll += DescViewFlow_Scroll;
                _descViewFlow.MouseWheel += DescViewFlow_MouseWheel;
            }
        }

        private void DescViewFlow_Scroll(object sender, ScrollEventArgs e)
        {
            if (_descriptionScrollBar == null || _descriptionScrollBar.IsDisposed) return;
            try { _descriptionScrollBar.ShowScrollBar(); } catch { }
        }

        private void DescViewFlow_MouseWheel(object sender, MouseEventArgs e)
        {
            if (_descriptionScrollBar == null || _descriptionScrollBar.IsDisposed) return;
            try { _descriptionScrollBar.ShowScrollBar(); } catch { }
        }

        protected override void OnLayoutDpiChanged()
        {
            base.OnLayoutDpiChanged();
            if (_wordListRowHeightImageList != null)
                _wordListRowHeightImageList.ImageSize = new Size(1, Math.Min(256, Px(38)));
            ResizeWordListColumn();
        }

        private void BuildEditPanel()
        {
            _editPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                AutoScroll = true,
                Visible = false,
                Padding = new Padding(0, ThemeManager.SpaceSM, 0, 0)
            };

            int y = ThemeManager.SpaceSM;
            int editW = 360;

            // Definition
            var lblDef = new EmojiLabel
            {
                Text = "\uD83D\uDCD6  \u041c\u0430\u044a\u043d\u043e\u0441\u0438",
                Font = _formFontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                Location = new Point(0, y),
                AutoSize = true
            };
            y += Math.Max(26, lblDef.Height + 6);

            _editDefinition = CreateEditTextBox(y, 60, editW);
            y += 68;

            // Spelling rule
            var lblSpell = new EmojiLabel
            {
                Text = "\u270F\uFE0F  \u0418\u043c\u043b\u043e \u049b\u043e\u0438\u0434\u0430\u0441\u0438",
                Font = _formFontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                Location = new Point(0, y),
                AutoSize = true
            };
            y += Math.Max(26, lblSpell.Height + 6);

            _editSpellingRule = CreateEditTextBox(y, 50, editW);
            y += 58;

            // Grammar note
            var lblGram = new EmojiLabel
            {
                Text = "\uD83D\uDCDD  \u0413\u0440\u0430\u043c\u043c\u0430\u0442\u0438\u043a\u0430",
                Font = _formFontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                Location = new Point(0, y),
                AutoSize = true
            };
            y += Math.Max(26, lblGram.Height + 6);

            _editGrammarNote = CreateEditTextBox(y, 50, editW);
            y += 58;

            // Examples
            var lblEx = new EmojiLabel
            {
                Text = "\uD83D\uDCA1  \u041c\u0438\u0441\u043e\u043b\u043b\u0430\u0440 (\u04b3\u0430\u0440 \u049b\u0430\u0442\u043e\u0440 \u2014 \u0431\u0438\u0440 \u043c\u0438\u0441\u043e\u043b)",
                Font = _formFontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                Location = new Point(0, y),
                AutoSize = true
            };
            y += Math.Max(26, lblEx.Height + 6);

            _editExamples = CreateEditTextBox(y, 70, editW);
            y += 78;

            // Save / Cancel buttons РІР‚вЂќ bigger
            _btnCancelEdit = new ModernButton
            {
                Text = "\u0411\u0435\u043a\u043e\u0440",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = _formFontLG,
                Size = new Size(100, 36),
                Location = new Point(0, y)
            };
            _btnCancelEdit.Click += BtnCancelEdit_Click;

            var btnSaveDesc = new ModernButton
            {
                Text = "\u0421\u0430\u049b\u043b\u0430\u0448",
                Style = ModernButton.ButtonStyle.Primary,
                Font = _formFontLG,
                Size = new Size(120, 36),
                Location = new Point(110, y)
            };
            btnSaveDesc.Click += BtnSaveDescription_Click;

            y += 46;

            _editPanel.Height = y;
            _editPanel.Controls.Add(lblDef);
            _editPanel.Controls.Add(_editDefinition);
            _editPanel.Controls.Add(lblSpell);
            _editPanel.Controls.Add(_editSpellingRule);
            _editPanel.Controls.Add(lblGram);
            _editPanel.Controls.Add(_editGrammarNote);
            _editPanel.Controls.Add(lblEx);
            _editPanel.Controls.Add(_editExamples);
            _editPanel.Controls.Add(_btnCancelEdit);
            _editPanel.Controls.Add(btnSaveDesc);
        }

        private TextBox CreateEditTextBox(int y, int height, int width)
        {
            return new TextBox
            {
                Location = new Point(0, y),
                Size = new Size(width, height),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = _formFontLG,
                BackColor = ThemeManager.Background,
                ForeColor = ThemeManager.TextPrimary,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private EmojiLabel CreateSectionLabel(string text)
        {
            return new EmojiLabel
            {
                Text = text,
                Font = _formFontLGBold,
                ForeColor = ThemeManager.TextPrimary,
                AutoSize = true,
                Padding = new Padding(0, ThemeManager.SpaceMD, 0, ThemeManager.SpaceXS),
                Visible = true
            };
        }

        private Label CreateSectionText()
        {
            return new Label
            {
                Text = "",
                Font = _formFontLG,
                ForeColor = ThemeManager.TextSecondary,
                AutoSize = true,
                MaximumSize = new Size(380, 0),
                Padding = new Padding(ThemeManager.SpaceSM, 0, 0, 0),
                Visible = true
            };
        }

        private TableLayoutPanel CreateSectionBlock(Control label, Label text)
        {
            var block = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 0, 0, ThemeManager.SpaceSM),
                Padding = new Padding(0),
                Visible = false
            };
            block.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            block.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            block.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            label.Dock = DockStyle.Fill;
            text.Dock = DockStyle.Fill;

            block.Controls.Add(label, 0, 0);
            block.Controls.Add(text, 0, 1);
            return block;
        }

        private void UpdateDescriptionViewWidths()
        {
            if (_descViewFlow == null) return;

            int available = Math.Max(220, _descViewFlow.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 10);
            _descNoDataLabel.MaximumSize = new Size(available, 0);

            int sectionTextWidth = Math.Max(200, available - ThemeManager.SpaceSM);
            _descDefinitionText.MaximumSize = new Size(sectionTextWidth, 0);
            _descSpellingText.MaximumSize = new Size(sectionTextWidth, 0);
            _descGrammarText.MaximumSize = new Size(sectionTextWidth, 0);
            _descExamplesText.MaximumSize = new Size(sectionTextWidth, 0);

            if (_descDefinitionSection != null) _descDefinitionSection.MaximumSize = new Size(available, 0);
            if (_descSpellingSection != null) _descSpellingSection.MaximumSize = new Size(available, 0);
            if (_descGrammarSection != null) _descGrammarSection.MaximumSize = new Size(available, 0);
            if (_descExamplesSection != null) _descExamplesSection.MaximumSize = new Size(available, 0);
        }

        private void UpdateDescription(string word)
        {
            // Exit edit mode when switching words
            if (_isEditMode) ExitEditMode();

            // Reset all sections
            _descDefinitionSection.Visible = false;
            _descSpellingSection.Visible = false;
            _descGrammarSection.Visible = false;
            _descExamplesSection.Visible = false;
            _descNoDataLabel.Visible = false;
            _btnEditSave.Visible = false;
            if (_descViewFlow != null) _descViewFlow.Visible = true;
            if (_editPanel != null) _editPanel.Visible = false;

            if (string.IsNullOrWhiteSpace(word))
            {
                _descWordLabel.Text = "\u0421\u045e\u0437\u043d\u0438 \u0442\u0430\u043d\u043b\u0430\u043d\u0433...";
                _descWordLabel.ForeColor = ThemeManager.TextSecondary;
                return;
            }

            _descWordLabel.Text = word;
            _descWordLabel.ForeColor = ThemeManager.Primary;
            _btnEditSave.Visible = (_explanationProvider != null);

            if (_explanationProvider == null)
            {
                _descNoDataLabel.Text = "\u0418\u0437\u043e\u04b3 \u0445\u0438\u0437\u043c\u0430\u0442\u0438 \u0443\u043b\u0430\u043d\u043c\u0430\u0433\u0430\u043d.";
                _descNoDataLabel.Visible = true;
                return;
            }

            // Resolve to Cyrillic form for explanation lookup (explanations are keyed by Cyrillic)
            string lookupWord = ResolveCyrillicWord(word);
            var entry = _explanationProvider.GetExplanation(lookupWord);
            if (entry == null)
            {
                _descNoDataLabel.Text = "\u0418\u0437\u043e\u04b3 \u043c\u0430\u0432\u0436\u0443\u0434 \u044d\u043c\u0430\u0441.\n\"\u0422\u0430\u04b3\u0440\u0438\u0440\" \u0442\u0443\u0433\u043c\u0430\u0441\u0438\u043d\u0438 \u0431\u043e\u0441\u0438\u0431 \u049b\u045e\u0448\u0438\u043d\u0433.";
                _descNoDataLabel.Visible = true;
                _btnEditSave.Text = "\u049a\u045e\u0448\u0438\u0448";
                return;
            }

            _btnEditSave.Text = "\u0422\u0430\u04b3\u0440\u0438\u0440";
            bool hasAny = false;

            // Definition
            if (!string.IsNullOrWhiteSpace(entry.Definition))
            {
                _descDefinitionText.Text = entry.Definition;
                _descDefinitionSection.Visible = true;
                hasAny = true;
            }

            // Spelling rule
            if (!string.IsNullOrWhiteSpace(entry.SpellingRule))
            {
                _descSpellingText.Text = entry.SpellingRule;
                _descSpellingSection.Visible = true;
                hasAny = true;
            }

            // Grammar note
            if (!string.IsNullOrWhiteSpace(entry.GrammarNote))
            {
                _descGrammarText.Text = entry.GrammarNote;
                _descGrammarSection.Visible = true;
                hasAny = true;
            }

            // Examples
            if (entry.Examples != null && entry.Examples.Length > 0)
            {
                _descExamplesText.Text = string.Join("\n", entry.Examples.Select(ex => "\u2022 " + ex));
                _descExamplesSection.Visible = true;
                hasAny = true;
            }

            if (!hasAny)
            {
                _descNoDataLabel.Text = "\u0418\u0437\u043e\u04b3 \u043c\u0430\u0432\u0436\u0443\u0434 \u044d\u043c\u0430\u0441.\n\"\u0422\u0430\u04b3\u0440\u0438\u0440\" \u0442\u0443\u0433\u043c\u0430\u0441\u0438\u043d\u0438 \u0431\u043e\u0441\u0438\u0431 \u049b\u045e\u0448\u0438\u043d\u0433.";
                _descNoDataLabel.Visible = true;
            }

            UpdateDescriptionViewWidths();
        }

        // =====================================================================
        //  EDIT MODE
        // =====================================================================

        private void EnterEditMode()
        {
            _isEditMode = true;
            string word = GetSelectedWord();
            if (string.IsNullOrWhiteSpace(word)) return;

            // Load existing data into edit fields (resolve to Cyrillic for lookup)
            string lookupWord = ResolveCyrillicWord(word);
            var entry = _explanationProvider?.GetExplanation(lookupWord);

            _editDefinition.Text = entry?.Definition ?? "";
            _editSpellingRule.Text = entry?.SpellingRule ?? "";
            _editGrammarNote.Text = entry?.GrammarNote ?? "";
            _editExamples.Text = entry?.Examples != null
                ? string.Join("\r\n", entry.Examples)
                : "";

            // Hide view sections, show edit panel.
            _descDefinitionSection.Visible = false;
            _descSpellingSection.Visible = false;
            _descGrammarSection.Visible = false;
            _descExamplesSection.Visible = false;
            _descNoDataLabel.Visible = false;
            if (_descViewFlow != null) _descViewFlow.Visible = false;
            if (_descriptionScrollBar != null) _descriptionScrollBar.Visible = false;

            _editPanel.Visible = true;
            _btnEditSave.Visible = false;

            _editDefinition.Focus();
        }

        private void ExitEditMode()
        {
            _isEditMode = false;
            _editPanel.Visible = false;
            if (_descViewFlow != null) _descViewFlow.Visible = true;
            if (_descriptionScrollBar != null)
            {
                _descriptionScrollBar.Visible = true;
            }
            _btnEditSave.Visible = true;
        }

        private void BtnEditSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(GetSelectedWord())) return;
            EnterEditMode();
        }

        private void BtnCancelEdit_Click(object sender, EventArgs e)
        {
            ExitEditMode();
            string word = GetSelectedWord();
            UpdateDescription(word);
        }

        private void BtnSaveDescription_Click(object sender, EventArgs e)
        {
            if (_explanationProvider == null) return;
            string word = GetSelectedWord();
            if (string.IsNullOrWhiteSpace(word)) return;

            // Resolve to Cyrillic for explanation storage
            string saveWord = ResolveCyrillicWord(word);

            // Parse examples: each line is one example, skip empty
            string[] examples = _editExamples.Text
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToArray();

            var entry = new ExplanationEntry
            {
                Word = saveWord,
                Definition = _editDefinition.Text.Trim(),
                SpellingRule = _editSpellingRule.Text.Trim(),
                GrammarNote = _editGrammarNote.Text.Trim(),
                Examples = examples.Length > 0 ? examples : null
            };

            // Check if at least one field has content
            bool hasContent = !string.IsNullOrWhiteSpace(entry.Definition) ||
                              !string.IsNullOrWhiteSpace(entry.SpellingRule) ||
                              !string.IsNullOrWhiteSpace(entry.GrammarNote) ||
                              (entry.Examples != null && entry.Examples.Length > 0);

            if (hasContent)
            {
                _explanationProvider.AddOrUpdate(entry);
                ToastNotification.Success("\u0418\u0437\u043e\u04b3 \u0441\u0430\u049b\u043b\u0430\u043d\u0434\u0438", word);
            }
            else
            {
                // All fields empty — remove the entry
                _explanationProvider.Remove(saveWord);
                ToastNotification.ShowInfo("\u0418\u0437\u043e\u04b3 \u045e\u0447\u0438\u0440\u0438\u043b\u0434\u0438", word);
            }

            ExitEditMode();
            UpdateDescription(word);
        }

        // =====================================================================
        //  DATA
        // =====================================================================

        private void RefreshWordList()
        {
            // Reload asynchronously — shows busy overlay while fetching and
            // transliterating ~183K words on a background System.Threading.Thread.
            QueueInitialWordLoad();
        }

        /// <summary>
        /// Pre-split _allWords into Cyrillic / Latin sublists once so pill
        /// switching is O(1) instead of rescanning the whole list each time.
        /// Cyrillic words are transliterated to Latin so the "Лотин" pill always
        /// shows the full word list in Latin script — even when the underlying
        /// user dictionary stores only Cyrillic canonical forms.
        /// </summary>
        private void PreSplitByScript()
        {
            var translit = ThisAddIn.Transliterator;

            var cyrList = new List<string>();
            var latList = new List<string>();
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < _allWords.Count; i++)
            {
                string w = _allWords[i];
                if (IsCyrillicWord(w))
                {
                    cyrList.Add(w);
                    // Generate a Latin equivalent for the Лотин pill
                    if (translit != null)
                    {
                        string lat = translit.ToLatin(w)?.ToLowerInvariant();
                        if (!string.IsNullOrEmpty(lat) && !map.ContainsKey(lat))
                        {
                            latList.Add(lat);
                            map[lat] = w; // reverse mapping: Latin display > Cyrillic actual
                        }
                    }
                }
                else
                {
                    // Already Latin — keep as-is
                    if (!map.ContainsKey(w))
                    {
                        latList.Add(w);
                        // No reverse mapping needed — it IS the actual word
                    }
                }
            }

            latList.Sort(UzbekStringComparer.Instance);
            _cyrWords = cyrList;
            _latWords = latList;
            _latToCyrMap = map;
        }

        /// <summary>
        /// Returns the canonical Cyrillic form for dictionary operations.
        /// When the "Лотин" pill is active, the displayed word is Latin but the
        /// dictionary stores the Cyrillic canonical form. This method resolves the
        /// mapping. For Cyrillic words or words not in the map, returns the input.
        /// </summary>
        private string ResolveCyrillicWord(string displayWord)
        {
            if (string.IsNullOrWhiteSpace(displayWord)) return displayWord;
            if (_latToCyrMap.TryGetValue(displayWord, out string cyr))
                return cyr;
            return displayWord;
        }

        private void QueueInitialWordLoad()
        {
            SetBusyState(true, "\u041b\u0443\u0493\u0430\u0442 \u044e\u043a\u043b\u0430\u043d\u043c\u043e\u049b\u0434\u0430...", "\u0418\u043b\u0442\u0438\u043c\u043e\u0441, \u043a\u0443\u0442\u0438\u043d\u0433...");
            // Run ALL heavy work (fetch + transliteration + sort) on background
            // System.Threading.Thread so the dialog paints instantly.
            Task.Run(() =>
            {
                var words = _getWords?.Invoke() ?? new List<string>();

                var translit = ThisAddIn.Transliterator;
                var cyrList = new List<string>(words.Count);
                var latList = new List<string>(words.Count);
                var map = new Dictionary<string, string>(words.Count, StringComparer.OrdinalIgnoreCase);

                for (int i = 0; i < words.Count; i++)
                {
                    string w = words[i];
                    if (IsCyrillicWord(w))
                    {
                        cyrList.Add(w);
                        if (translit != null)
                        {
                            string lat = translit.ToLatin(w)?.ToLowerInvariant();
                            if (!string.IsNullOrEmpty(lat) && !map.ContainsKey(lat))
                            {
                                latList.Add(lat);
                                map[lat] = w;
                            }
                        }
                    }
                    else
                    {
                        if (!map.ContainsKey(w))
                        {
                            latList.Add(w);
                        }
                    }
                }

                latList.Sort(UzbekStringComparer.Instance);
                return new DictionaryService.EditorWordCache
                {
                    AllWords = words,
                    CyrillicWords = cyrList,
                    LatinWords = latList,
                    LatinToCyrillicMap = map
                };
            })
            .ContinueWith(t =>
            {
                if (IsDisposed) return;
                if (t.IsFaulted)
                {
                    Logger.Error("QueueInitialWordLoad failed", t.Exception);
                    if (InvokeRequired)
                        try { Invoke(new Action(() => { if (!IsDisposed) SetBusyState(false); })); } catch { }
                    else
                        SetBusyState(false);
                    return;
                }
                var cache = t.Result;
                // Persist cache so next open is instant
                _onCacheBuilt?.Invoke(cache);

                // All UI-bound work must run on the UI System.Threading.Thread.
                // Marshal everything via Invoke to avoid race conditions.
                if (InvokeRequired)
                {
                    Invoke(new Action(() =>
                    {
                        if (IsDisposed) return;
                        _allWords = cache.AllWords;
                        _cyrWords = cache.CyrillicWords;
                        _latWords = cache.LatinWords;
                        _latToCyrMap = cache.LatinToCyrillicMap;
                        SetBusyState(false);
                        FilterWords();
                    }));
                }
                else
                {
                    _allWords = cache.AllWords;
                    _cyrWords = cache.CyrillicWords;
                    _latWords = cache.LatinWords;
                    _latToCyrMap = cache.LatinToCyrillicMap;
                    SetBusyState(false);
                    FilterWords();
                }
            }, TaskScheduler.Default);
        }

        // =================================================================
        //  PILL FILTER BUTTONS
        // =================================================================

        private static readonly string[] PillLabels = { "\u04B2\u0430\u043c\u043c\u0430\u0441\u0438", "\u041b\u043e\u0442\u0438\u043d", "\u041a\u0438\u0440\u0438\u043b\u043b" };

        private void BuildPills()
        {
            _pills.Clear();
            _pillBar.Controls.Clear();

            int x = 0;
            for (int i = 0; i < PillLabels.Length; i++)
            {
                var pill = new Label
                {
                    Text = PillLabels[i],
                    Font = i == 0 ? _fPill : _fPillR,
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    Tag = i
                };

                int tw = TextRenderer.MeasureText(PillLabels[i], _fPill).Width + 28;
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
            UpdatePillStyles();
        }

        private void SelectPill(int index)
        {
            _activePillIndex = index;
            UpdatePillStyles();
            FilterWords();
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

        private static bool IsCyrillicWord(string w)
        {
            foreach (char c in w)
                if (c >= '\u0400' && c <= '\u04FF') return true;
            return false;
        }

        private static bool IsLatinWord(string w)
        {
            foreach (char c in w)
                if (c >= '\u0400' && c <= '\u04FF') return false;
            return true;
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

        private void FilterWords()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(FilterWords));
                return;
            }

            if (_wordList == null || _countLabel == null) return;

            string query = _searchBox?.Text?.Trim() ?? "";

            // Pill filter — use pre-split lists (O(1) switch, no per-word scan)
            List<string> source;
            if      (_activePillIndex == 1) source = _latWords;   // Лотин
            else if (_activePillIndex == 2) source = _cyrWords;   // Кирилл
            else                            source = _allWords;   // ?аммаси

            if (string.IsNullOrEmpty(query))
            {
                _filteredWords = source;
            }
            else
            {
                _filteredWords = source
                    .Where(w => !string.IsNullOrEmpty(w) &&
                                w.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }

            BindAllAtOnce();
        }

        private void BindAllAtOnce()
        {
            string previouslySelected = GetSelectedWord();

            _wordList.BeginUpdate();
            try
            {
                _wordList.VirtualListSize = 0;
                _wordList.VirtualListSize = _filteredWords.Count;
                ResizeWordListColumn();
            }
            finally
            {
                _wordList.EndUpdate();
            }

            _wordList.SelectedIndices.Clear();
            if (!string.IsNullOrEmpty(previouslySelected))
            {
                int idx = _filteredWords.FindIndex(w => string.Equals(w, previouslySelected, StringComparison.OrdinalIgnoreCase));
                if (idx >= 0 && idx < _wordList.VirtualListSize)
                {
                    _wordList.Items[idx].Selected = true;
                    _wordList.EnsureVisible(idx);
                }
            }

            _countLabel.Text = $"{_filteredWords.Count} \u0441\u045e\u0437";
        }

        private void WordList_RetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e)
        {
            if (e.ItemIndex < 0 || e.ItemIndex >= _filteredWords.Count)
            {
                e.Item = new ListViewItem(string.Empty);
                return;
            }

            e.Item = new ListViewItem(_filteredWords[e.ItemIndex]) { ImageIndex = 0 };
        }

        private void ResizeWordListColumn()
        {
            if (_wordList == null || _wordList.Columns.Count == 0) return;
            // Reserve room for vertical scrollbar to avoid horizontal scrollbar flicker/appearance.
            int width = Math.Max(120, _wordList.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 8);
            _wordList.Columns[0].Width = width;
        }

        private string GetSelectedWord()
        {
            if (_wordList == null || _wordList.SelectedIndices.Count == 0)
                return null;

            int idx = _wordList.SelectedIndices[0];
            if (idx < 0 || idx >= _filteredWords.Count)
                return null;

            return _filteredWords[idx];
        }

        // =====================================================================
        //  EVENTS
        // =====================================================================

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (_isBusy) return;

            string word = _newWordBox.Text?.Trim();
            if (string.IsNullOrEmpty(word))
            {
                SafeExecutor.ShowWarning("\u0421\u045e\u0437 \u043a\u0438\u0440\u0438\u0442\u0438\u043d\u0433.");
                return;
            }
            // Check if already in dictionary before adding
            var dict = ThisAddIn.DictionaryService;
            if (dict != null && dict.Contains(word))
            {
                ToastNotification.ShowInfo("\u041b\u0443\u0493\u0430\u0442\u0434\u0430 \u043c\u0430\u0432\u0436\u0443\u0434", $"\"{word}\" \u0430\u043b\u043b\u0430\u049b\u0430\u0447\u043e\u043d \u043b\u0443\u0493\u0430\u0442\u0434\u0430 \u0431\u043e\u0440.");
                return;
            }
            _onAddWord?.Invoke(word);
            _newWordBox.Text = "";
            RefreshWordList();
            ToastNotification.Success("\u049a\u045e\u0448\u0438\u043b\u0434\u0438", word);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_isBusy) return;

            string word = GetSelectedWord();
            if (string.IsNullOrWhiteSpace(word)) return;

            // Resolve to Cyrillic canonical form for the actual removal
            string removeWord = ResolveCyrillicWord(word);

            // Prevent deletion of words from the main (built-in) dictionary
            if (_isMainWord != null && _isMainWord(removeWord))
            {
                ToastNotification.ShowInfo(
                    "\u0410\u0441\u043e\u0441\u0438\u0439 \u043b\u0443\u0493\u0430\u0442",
                    $"\"{word}\" \u0430\u0441\u043e\u0441\u0438\u0439 \u043b\u0443\u0493\u0430\u0442\u0434\u0430 \u0431\u045e\u043b\u0433\u0430\u043d\u0438 \u0443\u0447\u0443\u043d \u045e\u0447\u0438\u0440\u0438\u0431 \u0431\u045e\u043b\u043c\u0430\u0439\u0434\u0438.");
                return;
            }

            if (SafeExecutor.Confirm($"\"{word}\" \u0441\u045e\u0437\u0438\u043d\u0438 \u045e\u0447\u0438\u0440\u043c\u043e\u049b\u0447\u0438\u043c\u0438\u0441\u0438\u0437?"))
            {
                _onRemoveWord?.Invoke(removeWord);
                RefreshWordList();
            }
        }

        private void WordList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                BtnDelete_Click(sender, e);
                e.Handled = true;
            }
        }

        private void WordList_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedWord = GetSelectedWord();
            UpdateDescription(selectedWord);
        }

        private async void BtnExport_Click(object sender, EventArgs e)
        {
            if (_isBusy) return;

            var dict = ThisAddIn.DictionaryService;
            if (dict == null)
            {
                SafeExecutor.ShowWarning("\u041b\u0443\u0493\u0430\u0442 \u0445\u0438\u0437\u043c\u0430\u0442\u0438 \u044e\u043a\u043b\u0430\u043d\u043c\u0430\u0433\u0430\u043d.");
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "\u041b\u0443\u0493\u0430\u0442\u043d\u0438 \u044d\u043a\u0441\u043f\u043e\u0440\u0442 \u049b\u0438\u043b\u0438\u0448";
                dialog.Filter =
                    "JSON \u0444\u0430\u0439\u043b — \u0442\u045e\u043b\u0438\u049b \u0437\u0430\u0445\u0438\u0440\u0430 (*.json)|*.json|" +
                    "\u041b\u0443\u0493\u0430\u0442 \u0440\u045e\u0439\u0445\u0430\u0442\u0438 (*.dic)|*.dic|" +
                    "Excel \u0444\u0430\u0439\u043b (*.xlsx)|*.xlsx|Excel 97-2003 \u0444\u0430\u0439\u043b (*.xls)|*.xls";
                dialog.DefaultExt = "json";
                dialog.AddExtension = true;
                dialog.FileName = $"uzbekorfo_dictionary_{DateTime.Now:yyyyMMdd_HHmmss}";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                try
                {
                    string targetPath = dialog.FileName;
                    string ext = Path.GetExtension(targetPath).ToLowerInvariant();
                    string detail = "Файл тайёрланмо?да";
                    if (ext == ".json") detail = "JSON файл тайёрланмо?да";
                    else if (ext == ".dic") detail = "DIC \u0441\u045e\u0437\u043b\u0430\u0440 \u0440\u045e\u0439\u0445\u0430\u0442\u0438 \u0442\u0430\u0439\u0451\u0440\u043b\u0430\u043d\u043c\u043e\u049b\u0434\u0430";
                    else if (ext == ".xlsx") detail = "Excel (.xlsx) файл тайёрланмо?да";
                    else if (ext == ".xls") detail = "Excel 97-2003 (.xls) файл тайёрланмо?да";

                    SetBusyState(
                        true,
                        "\u041b\u0443\u0493\u0430\u0442 \u044d\u043a\u0441\u043f\u043e\u0440\u0442 \u049b\u0438\u043b\u0438\u043d\u043c\u043e\u049b\u0434\u0430...",
                        detail);
                    var provider = _explanationProvider;

                    var result = await Task.Run(() =>
                    {
                        var explanations = provider != null ? provider.GetAllEntries() : null;
                        return dict.ExportForMigration(targetPath, explanations);
                    });

                    string message = $"{result.WordCount} \u0442\u0430 \u0441\u045e\u0437 \u044d\u043a\u0441\u043f\u043e\u0440\u0442 \u049b\u0438\u043b\u0438\u043d\u0434\u0438.";
                    if (result.DefinitionCount > 0)
                        message += $" {result.DefinitionCount} \u0442\u0430 \u0438\u0437\u043e\u04b3 \u049b\u045e\u0448\u0438\u043b\u0434\u0438.";

                    ToastNotification.Success("\u042d\u043a\u0441\u043f\u043e\u0440\u0442 \u0442\u0443\u0433\u0430\u0434\u0438", message);
                }
                catch (Exception ex)
                {
                    Logger.Error("\u041b\u0443\u0493\u0430\u0442 \u044d\u043a\u0441\u043f\u043e\u0440\u0442\u0438\u0434\u0430 \u0445\u0430\u0442\u043e", ex);
                    string message =
                        ex is FileNotFoundException ||
                        ex is UnauthorizedAccessException ||
                        ex is IOException ||
                        ex is NotSupportedException ||
                        ex is InvalidOperationException ||
                        ex is COMException
                            ? ex.Message
                            : "\u042d\u043a\u0441\u043f\u043e\u0440\u0442 \u043f\u0430\u0439\u0442\u0438\u0434\u0430 \u0445\u0430\u0442\u043e \u044e\u0437 \u0431\u0435\u0440\u0434\u0438.";
                    SafeExecutor.ShowWarning(message);
                }
                finally
                {
                    SetBusyState(false);
                }
            }
        }

        private async void BtnImport_Click(object sender, EventArgs e)
        {
            if (_isBusy) return;

            var dict = ThisAddIn.DictionaryService;
            if (dict == null)
            {
                SafeExecutor.ShowWarning("\u041b\u0443\u0493\u0430\u0442 \u0445\u0438\u0437\u043c\u0430\u0442\u0438 \u044e\u043a\u043b\u0430\u043d\u043c\u0430\u0433\u0430\u043d.");
                return;
            }

            using (var dialog = ImportDictionaryWorkflowService.BuildImportDialog())
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                try
                {
                    SetBusyState(
                        true,
                        "\u041b\u0443\u0493\u0430\u0442 \u0438\u043c\u043f\u043e\u0440\u0442 \u049b\u0438\u043b\u0438\u043d\u043c\u043e\u049b\u0434\u0430...",
                        "\u0424\u0430\u0439\u043b \u0442\u0430\u04b3\u043b\u0438\u043b \u049b\u0438\u043b\u0438\u043d\u043c\u043e\u049b\u0434\u0430");

                    string sourcePath = dialog.FileName;
                    var provider = _explanationProvider;

                    var import = await RunStaAsync(() => dict.ImportFromFileDetailed(sourcePath));

                    int defsImported = 0;
                    if (provider != null && import.DefinitionEntries.Count > 0)
                    {
                        defsImported = await Task.Run(() =>
                            provider.AddOrUpdateBatch(import.DefinitionEntries));
                    }

                    RefreshWordList();
                    UpdateDescription(GetSelectedWord());

                    if (import.AddedWordCount == 0 && defsImported == 0)
                    {
                        ToastNotification.ShowInfo("\u0418\u043c\u043f\u043e\u0440\u0442 \u0442\u0443\u0433\u0430\u0434\u0438", "\u042f\u043d\u0433\u0438 \u0441\u045e\u0437 \u0451\u043a\u0438 \u0438\u0437\u043e\u04b3 \u0442\u043e\u043f\u0438\u043b\u043c\u0430\u0434\u0438.");
                    }
                    else
                    {
                        string msg = $"{import.AddedWordCount} \u0442\u0430 \u044f\u043d\u0433\u0438 \u0441\u045e\u0437 \u049b\u045e\u0448\u0438\u043b\u0434\u0438.";
                        if (defsImported > 0)
                            msg += $" {defsImported} \u0442\u0430 \u0438\u0437\u043e\u04b3 \u044e\u043a\u043b\u0430\u043d\u0434\u0438.";
                        ToastNotification.Success("\u0418\u043c\u043f\u043e\u0440\u0442 \u0442\u0443\u0433\u0430\u0434\u0438", msg);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("Лу?ат импортида хато", ex);
                    string message =
                        ex is FileNotFoundException ||
                        ex is UnauthorizedAccessException ||
                        ex is IOException ||
                        ex is NotSupportedException ||
                        ex is InvalidOperationException ||
                        ex is COMException
                            ? ex.Message
                            : "Импорт пайтида хато юз берди.";
                    SafeExecutor.ShowWarning(message);
                }
                finally
                {
                    SetBusyState(false);
                }
            }
        }

        private static Task<T> RunStaAsync<T>(Func<T> work)
        {
            if (work == null) throw new ArgumentNullException(nameof(work));

            var tcs = new TaskCompletionSource<T>();
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    var result = work();
                    tcs.TrySetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            });

            thread.IsBackground = true;
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }

        private void SetBusyState(bool isBusy, string title = null, string detail = null)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => SetBusyState(isBusy, title, detail)));
                return;
            }

            _isBusy = isBusy;

            if (_btnExportAction != null) _btnExportAction.Enabled = !isBusy;
            if (_btnImportAction != null) _btnImportAction.Enabled = !isBusy;

            if (_busyOverlay != null)
            {
                if (isBusy)
                {
                    _busyTitleLabel.Text = string.IsNullOrWhiteSpace(title)
                        ? "\u0418\u0448\u043b\u0430\u043d\u043c\u043e\u049b\u0434\u0430..."
                        : title;
                    _busyDetailLabel.Text = string.IsNullOrWhiteSpace(detail)
                        ? "\u0418\u043b\u0442\u0438\u043c\u043e\u0441, \u043a\u0443\u0442\u0438\u043d\u0433..."
                        : detail;
                    // Start marquee animation only while overlay is visible
                    if (_busyProgress != null) _busyProgress.Style = ProgressBarStyle.Marquee;
                    _busyOverlay.Visible = true;
                    _busyOverlay.BringToFront();
                }
                else
                {
                    _busyOverlay.Visible = false;
                    // Stop marquee animation to avoid background CPU usage
                    if (_busyProgress != null) _busyProgress.Style = ProgressBarStyle.Blocks;
                }
            }

            UseWaitCursor = isBusy;
            Cursor.Current = isBusy ? Cursors.WaitCursor : Cursors.Default;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_searchDebounce != null)
                {
                    _searchDebounce.Stop();
                    _searchDebounce.Dispose();
                    _searchDebounce = null;
                }

                if (_wordListRowHeightImageList != null)
                {
                    _wordListRowHeightImageList.Dispose();
                    _wordListRowHeightImageList = null;
                }

                if (_wordListScrollBar != null)
                {
                    _wordListScrollBar.Dispose();
                    _wordListScrollBar = null;
                }

                if (_descriptionScrollBar != null)
                {
                    _descriptionScrollBar.Dispose();
                    _descriptionScrollBar = null;
                }

                if (_descViewFlow != null)
                {
                    _descViewFlow.Scroll -= DescViewFlow_Scroll;
                    _descViewFlow.MouseWheel -= DescViewFlow_MouseWheel;
                }
            }

            base.Dispose(disposing);
        }
    }
}











