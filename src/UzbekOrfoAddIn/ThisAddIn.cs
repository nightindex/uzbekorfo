using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using UzbekOrfoAddIn.Services;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn
{
    public partial class ThisAddIn
    {
        // =====================================================================
        //  GLOBAL SERVICE INSTANCES — accessible from all ribbon handlers
        // =====================================================================

        /// <summary>Application settings and file paths.</summary>
        public static SettingsManager Settings { get; private set; }

        /// <summary>In-memory store of errors from the last spelling check.</summary>
        public static ErrorStore ErrorStore { get; private set; }

        /// <summary>Dictionary service for word look-up and management.</summary>
        public static DictionaryService DictionaryService { get; private set; }

        /// <summary>Spelling engine for checking text and generating suggestions.</summary>
        public static SpellingEngine SpellingEngine { get; private set; }

        /// <summary>Real-time auto-correct service.</summary>
        public static AutoCorrectService AutoCorrectService { get; private set; }

        /// <summary>Transliteration service for Latin ↔ Cyrillic conversion.</summary>
        public static TransliterationService Transliterator { get; private set; }

        /// <summary>Explanation/definition provider for Uzbek words.</summary>
        public static ExplanationProvider ExplanationProvider { get; private set; }

        /// <summary>Morphological analyzer for Uzbek words.</summary>
        public static UzbekMorphAnalyzer MorphAnalyzer { get; private set; }

        /// <summary>Grammar checking engine.</summary>
        public static GrammarEngine GrammarEngine { get; private set; }

        // Future services (will be initialized as implemented):
        // public static PredictionEngine PredictionEngine { get; private set; }

        private static bool _globalExceptionHandlersRegistered;
        private AddInRuntime _runtime;

        /// <summary>
        /// Indicates whether all core services were initialized successfully.
        /// When false, ribbon/context-menu actions must not proceed.
        /// </summary>
        public static bool IsInitialized { get; private set; }

        // =====================================================================
        //  STARTUP / SHUTDOWN
        // =====================================================================

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Logger.Info("=== Ўзбек Орфо Add-In бошланмоқда ===");
                RegisterGlobalExceptionHandlers();
                InitializeRuntime();
                IsInitialized = true;

                // Clean up any orphaned context menu buttons from previous sessions
                CleanAllContextMenuButtons();

                // Hook right-click context menu
                this.Application.WindowBeforeRightClick +=
                    new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);

                // Keep spell highlights non-persistent.
                this.Application.DocumentOpen +=
                    new Word.ApplicationEvents4_DocumentOpenEventHandler(App_DocumentOpen);
                this.Application.DocumentBeforeSave +=
                    new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(App_DocumentBeforeSave);
                this.Application.DocumentBeforeClose +=
                    new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(App_DocumentBeforeClose);

                // Missing-file preflight checks disabled by request.

                // Pre-warm the dictionary editor form shell on the UI thread
                // after a short delay so Word finishes loading first.
                // Words load asynchronously when the form is first shown.
                var preWarmTimer = new System.Windows.Forms.Timer { Interval = 500 };
                preWarmTimer.Tick += (ts, te) =>
                {
                    preWarmTimer.Stop();
                    preWarmTimer.Dispose();
                    try
                    {
                        EditDictionaryWorkflowService.PreWarmForm(DictionaryService, ExplanationProvider);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn($"Form pre-warm failed (non-fatal): {ex.Message}");
                    }
                };
                preWarmTimer.Start();
            }
            catch (Exception ex)
            {
                Logger.Error("Startup failed", ex);
                IsInitialized = false;
                try
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Ўзбек Орфо қўшимчаси юкланишда хато юз берди.\nТафсилотлар: " + ex.Message,
                        "Ўзбек Орфо",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }
                catch { }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                try { _runtime?.Shutdown(); }
                catch (Exception ex) { Logger.Warn($"Runtime shutdown error: {ex.Message}"); }

                // Remove keyboard hotkeys
                try { HotkeyManager.Unregister(); }
                catch (Exception ex) { Logger.Warn($"HotkeyManager unregister error: {ex.Message}"); }

                // Clean up context menu
                RemoveContextMenuControls();

                // Unhook application events
                try
                {
                    this.Application.WindowBeforeRightClick -=
                        new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);
                    this.Application.DocumentOpen -=
                        new Word.ApplicationEvents4_DocumentOpenEventHandler(App_DocumentOpen);
                    this.Application.DocumentBeforeSave -=
                        new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(App_DocumentBeforeSave);
                    this.Application.DocumentBeforeClose -=
                        new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(App_DocumentBeforeClose);
                }
                catch (Exception ex) { Logger.Warn($"Event unhook error: {ex.Message}"); }

                _runtime = null;
                Settings = null;
                ErrorStore = null;
                DictionaryService = null;
                SpellingEngine = null;
                AutoCorrectService = null;
                Transliterator = null;
                ExplanationProvider = null;
                MorphAnalyzer = null;
                GrammarEngine = null;

                IsInitialized = false;
                Logger.Info("=== Ўзбек Орфо Add-In тўхтатилди ===");
                Logger.Flush(); // ensure all buffered log lines are written before exit
            }
            catch (Exception ex)
            {
                Logger.Error("Shutdown error", ex);
            }
        }

        private void InitializeRuntime()
        {
            _runtime = AddInRuntime.Initialize(this.Application);

            Settings = _runtime.Settings;
            ErrorStore = _runtime.ErrorStore;
            DictionaryService = _runtime.DictionaryService;
            SpellingEngine = _runtime.SpellingEngine;
            AutoCorrectService = _runtime.AutoCorrectService;
            Transliterator = _runtime.Transliterator;
            ExplanationProvider = _runtime.ExplanationProvider;
            MorphAnalyzer = _runtime.MorphAnalyzer;
            GrammarEngine = _runtime.GrammarEngine;
        }

        private static void RegisterGlobalExceptionHandlers()
        {
            if (_globalExceptionHandlersRegistered) return;
            _globalExceptionHandlersRegistered = true;

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                try
                {
                    var ex = e.ExceptionObject as Exception;
                    if (ex != null) Logger.Error("Unhandled domain exception", ex);
                    else Logger.Error($"Unhandled domain exception: {e.ExceptionObject}");
                }
                catch { }
            };

            System.Threading.Tasks.TaskScheduler.UnobservedTaskException += (s, e) =>
            {
                try { Logger.Error("Unobserved task exception", e.Exception); }
                catch { }
                finally { e.SetObserved(); }
            };

            try
            {
                System.Windows.Forms.Application.ThreadException += (s, e) =>
                {
                    try { Logger.Error("WinForms UI thread exception", e.Exception); }
                    catch { }
                };
            }
            catch { }
        }

        private void RunStartupPreflightChecks()
        {
            // Missing-file startup checks intentionally disabled.
            // Seeding and service initialization continue to handle recovery paths.
        }

        // =====================================================================
        //  RIGHT-CLICK CONTEXT MENU
        // =====================================================================
        //
        //  CommandBar-based approach: Word’s Fluent UI merges legacy CommandBar
        //  controls into the context menu. We inject a popup ("Имло") with
        //  spelling suggestions, or a standalone "Луғатга қўшиш" button.
        //  All controls are tagged CTX_TAG and created with Temporary=true
        //  so cleanup is simple and reliable.
        //
        // =====================================================================

        private const string CTX_TAG = "UzbekOrfo_Ctx";
        private const string CTX_POPUP_CAPTION = "Имло";
        private const int CTX_MAX_SUGGESTIONS = 8;
        private const int CTX_SPELLING_CACHE_MAX = 256;
        private const int CTX_GRAMMAR_CACHE_MAX = 128;
        private const string SPELL_HIGHLIGHT_FLAG_VAR = "UzbekOrfo_HasSpellHighlights";
        private const string SPELL_HIGHLIGHT_MIGRATED_VAR = "UzbekOrfo_SpellHighlightMigrated";

        // ---- Static bridging properties for Ribbon XML context menu ----

        /// <summary>Whether the word under the cursor is misspelled (used by ContextMenuRibbon / CombinedRibbon).</summary>
        public static bool CtxMenuIsMisspelled
        {
            get
            {
                try
                {
                    var addIn = Globals.ThisAddIn;
                    if (addIn == null || string.IsNullOrWhiteSpace(addIn._ctxTargetWord)) return false;
                    addIn.GetContextSpellingData(addIn._ctxTargetWord, CTX_MAX_SUGGESTIONS, out bool isMisspelled, out _);
                    return isMisspelled;
                }
                catch { return false; }
            }
        }

        /// <summary>Spelling suggestions for the context-menu target word.</summary>
        public static List<string> CtxMenuSuggestions
        {
            get
            {
                try
                {
                    var addIn = Globals.ThisAddIn;
                    if (addIn == null || string.IsNullOrWhiteSpace(addIn._ctxTargetWord)) return new List<string>();
                    addIn.GetContextSpellingData(addIn._ctxTargetWord, CTX_MAX_SUGGESTIONS, out _, out List<string> suggestions);
                    return suggestions;
                }
                catch { return new List<string>(); }
            }
        }

        /// <summary>The raw word captured on right-click for the context menu.</summary>
        public static string CtxMenuRawWord
        {
            get
            {
                try { return Globals.ThisAddIn?._ctxTargetWord; }
                catch { return null; }
            }
        }

        /// <summary>Replace the context-menu target word with the given replacement text.</summary>
        public static void ReplaceContextTargetWord(string replacement)
        {
            try
            {
                var addIn = Globals.ThisAddIn;
                if (addIn == null || string.IsNullOrWhiteSpace(replacement)) return;

                var range = addIn.GetContextMenuTargetRange();
                if (range == null) return;

                string oldWord = TextHelper.NormalizeWord(range.Text ?? addIn._ctxTargetWord ?? string.Empty);

                DocumentHelper.BeginUndoRecord("Сўзни тузатиш");
                try { DocumentHelper.ReplaceRangeText(range, replacement); }
                finally { DocumentHelper.EndUndoRecord(); }

                TryClearSelectionUnderline();
                if (!string.IsNullOrWhiteSpace(oldWord))
                    ResolveErrorsForWord(oldWord);
            }
            catch (Exception ex)
            {
                Logger.Error("ReplaceContextTargetWord error", ex);
            }
        }

        private readonly List<Office.CommandBarControl> _ctxControls = new List<Office.CommandBarControl>();
        private readonly object _ctxSpellingCacheLock = new object();
        private readonly object _ctxGrammarCacheLock = new object();
        private readonly Dictionary<string, CachedContextSpelling> _ctxSpellingCache =
            new Dictionary<string, CachedContextSpelling>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, CachedContextGrammar> _ctxGrammarCache =
            new Dictionary<string, CachedContextGrammar>(StringComparer.Ordinal);
        private static readonly TimeSpan _ctxSpellingCacheTtl = TimeSpan.FromMinutes(8);
        private static readonly TimeSpan _ctxGrammarCacheTtl = TimeSpan.FromMinutes(2);

        private string _ctxTargetDocKey;
        private int _ctxTargetWordStart = -1;
        private int _ctxTargetWordEnd = -1;
        private string _ctxTargetWord;

        private static readonly object _addWordRequestLock = new object();
        private static string _lastAddRequestWord;
        private static DateTime _lastAddRequestAtUtc = DateTime.MinValue;
        private static Models.AddWordResult _lastAddRequestResult = Models.AddWordResult.Invalid;
        private static readonly TimeSpan _duplicateAddRequestWindow = TimeSpan.FromMilliseconds(1200);

        private sealed class CachedContextSpelling
        {
            public bool IsMisspelled { get; set; }
            public List<string> Suggestions { get; set; } = new List<string>();
            public DateTime CachedAtUtc { get; set; }
        }

        private sealed class CachedContextGrammar
        {
            public List<Models.ErrorEntry> Errors { get; set; } = new List<Models.ErrorEntry>();
            public DateTime CachedAtUtc { get; set; }
        }

        // ---------------------------------------------------------------------
        //  Cleanup
        // ---------------------------------------------------------------------

        /// <summary>Startup cleanup: remove any orphaned controls from previous sessions.</summary>
        private void CleanAllContextMenuButtons()
        {
            try { DeepCleanContextBar("Text"); } catch { }
            try { DeepCleanContextBar("Spelling"); } catch { }
        }

        /// <summary>Delete every control with our tag or orphaned ghost items from the named CommandBar.</summary>
        private void DeepCleanContextBar(string menuName)
        {
            try
            {
                var bar = this.Application.CommandBars[menuName];
                if (bar == null) return;

                // Reverse-iterate; restart after each delete (COM index shifts).
                bool deleted = true;
                int safety = 0;
                while (deleted && safety++ < 100)
                {
                    deleted = false;
                    for (int i = bar.Controls.Count; i >= 1; i--)
                    {
                        try
                        {
                            var ctl = bar.Controls[i];
                            bool isOurs = (ctl.Tag ?? string.Empty).StartsWith(CTX_TAG, StringComparison.Ordinal);

                            if (!isOurs)
                            {
                                bool isBuiltIn = true;
                                try { isBuiltIn = ctl.BuiltIn; } catch { }

                                if (!isBuiltIn)
                                {
                                    string cap = ctl.Caption?.Replace("&", "")?.Trim() ?? "";
                                    if (cap == CTX_POPUP_CAPTION || 
                                        cap.StartsWith("Луғатга қўшиш") || 
                                        string.IsNullOrWhiteSpace(cap))
                                    {
                                        isOurs = true;
                                    }
                                }
                            }

                            if (isOurs)
                            {
                                ctl.Delete(false);
                                deleted = true;
                                break;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        /// <summary>Per right-click cleanup: delete tracked controls then sweep by tag.</summary>
        private void RemoveContextMenuControls()
        {
            // Reverse delete to prevent leaving orphaned children in the Office Ribbon cache
            for (int i = _ctxControls.Count - 1; i >= 0; i--)
            {
                try { _ctxControls[i]?.Delete(false); } catch { }
            }
            _ctxControls.Clear();

            DeepCleanContextBar("Text");
            DeepCleanContextBar("Spelling");
            ResetContextMenuTarget();
        }

        // ---------------------------------------------------------------------
        //  Right-Click Handler
        // ---------------------------------------------------------------------

        private void App_WindowBeforeRightClick(Word.Selection sel, ref bool cancel)
        {
            try
            {
                RemoveContextMenuControls();

                string rawWord = null;
                string normalizedWord = null;
                CaptureContextMenuTarget(sel, out rawWord, out normalizedWord);

                // Notify Ribbon UI that the context word has changed so it can redraw XML menus
                ContextMenuRibbon.InvalidateContextControls();

                if (string.IsNullOrWhiteSpace(normalizedWord) || normalizedWord.Length < 2)
                    return;

                string addLabel = rawWord != null && rawWord.Length >= 2
                    ? "Луғатга қўшиш:  \"" + rawWord + "\""
                    : "Луғатга қўшиш";

                if (SpellingEngine != null)
                {
                    int maxSuggestions = 5;
                    try
                    {
                        if (Settings != null)
                            maxSuggestions = Math.Max(1, Math.Min(CTX_MAX_SUGGESTIONS, Settings.MaxSpellingSuggestions));
                    }
                    catch (Exception ex) { Logger.Warn($"MaxSpellingSuggestions read error: {ex.Message}"); }

                    bool isMisspelled;
                    var suggestions = new List<string>();
                    GetContextSpellingData(normalizedWord, maxSuggestions, out isMisspelled, out suggestions);

                    if (isMisspelled)
                        InsertSpellingSuggestions(addLabel, suggestions);
                    else
                        InsertAddToDictButton(addLabel);
                }
                else
                {
                    InsertAddToDictButton(addLabel);
                }

                // Grammar suggestions in context menu are disabled to avoid
                // intrusive suggestion entries on simple word clicks.
            }
            catch (Exception ex)
            {
                Logger.Error("Context menu error", ex);
            }
        }

        private void GetContextSpellingData(
            string normalizedWord,
            int maxSuggestions,
            out bool isMisspelled,
            out List<string> suggestions)
        {
            isMisspelled = false;
            suggestions = new List<string>();
            if (string.IsNullOrWhiteSpace(normalizedWord) || SpellingEngine == null)
                return;

            int boundedMax = Math.Max(1, Math.Min(CTX_MAX_SUGGESTIONS, maxSuggestions));
            var now = DateTime.UtcNow;

            lock (_ctxSpellingCacheLock)
            {
                if (_ctxSpellingCache.TryGetValue(normalizedWord, out var cached) &&
                    (now - cached.CachedAtUtc) <= _ctxSpellingCacheTtl)
                {
                    isMisspelled = cached.IsMisspelled;
                    if (isMisspelled && cached.Suggestions != null)
                        suggestions = cached.Suggestions.Take(boundedMax).ToList();
                    return;
                }
            }

            bool computedMisspelled = !SpellingEngine.IsCorrect(normalizedWord);
            var computedSuggestions = new List<string>();

            if (computedMisspelled)
            {
                computedSuggestions = SpellingEngine
                    .GetSuggestions(normalizedWord, CTX_MAX_SUGGESTIONS)
                    .Select(s => s?.Text)
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .Where(t => !string.Equals(TextHelper.NormalizeWord(t), normalizedWord, StringComparison.OrdinalIgnoreCase))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Take(CTX_MAX_SUGGESTIONS)
                    .ToList();
            }

            lock (_ctxSpellingCacheLock)
            {
                _ctxSpellingCache[normalizedWord] = new CachedContextSpelling
                {
                    IsMisspelled = computedMisspelled,
                    Suggestions = computedSuggestions,
                    CachedAtUtc = now
                };

                if (_ctxSpellingCache.Count > CTX_SPELLING_CACHE_MAX)
                {
                    int removeCount = _ctxSpellingCache.Count - CTX_SPELLING_CACHE_MAX;
                    foreach (var key in _ctxSpellingCache
                        .OrderBy(kv => kv.Value.CachedAtUtc)
                        .Take(removeCount)
                        .Select(kv => kv.Key)
                        .ToList())
                    {
                        _ctxSpellingCache.Remove(key);
                    }
                }
            }

            isMisspelled = computedMisspelled;
            if (computedMisspelled)
                suggestions = computedSuggestions.Take(boundedMax).ToList();
        }

        private List<Models.ErrorEntry> GetContextGrammarErrors(string sentenceText)
        {
            if (string.IsNullOrWhiteSpace(sentenceText) || GrammarEngine == null)
                return new List<Models.ErrorEntry>();

            string key = sentenceText;
            var now = DateTime.UtcNow;

            lock (_ctxGrammarCacheLock)
            {
                if (_ctxGrammarCache.TryGetValue(key, out var cached) &&
                    (now - cached.CachedAtUtc) <= _ctxGrammarCacheTtl)
                {
                    return cached.Errors != null
                        ? new List<Models.ErrorEntry>(cached.Errors)
                        : new List<Models.ErrorEntry>();
                }
            }

            var computed = GrammarEngine.CheckText(sentenceText) ?? new List<Models.ErrorEntry>();

            lock (_ctxGrammarCacheLock)
            {
                _ctxGrammarCache[key] = new CachedContextGrammar
                {
                    Errors = new List<Models.ErrorEntry>(computed),
                    CachedAtUtc = now
                };

                if (_ctxGrammarCache.Count > CTX_GRAMMAR_CACHE_MAX)
                {
                    int removeCount = _ctxGrammarCache.Count - CTX_GRAMMAR_CACHE_MAX;
                    foreach (var oldKey in _ctxGrammarCache
                        .OrderBy(kv => kv.Value.CachedAtUtc)
                        .Take(removeCount)
                        .Select(kv => kv.Key)
                        .ToList())
                    {
                        _ctxGrammarCache.Remove(oldKey);
                    }
                }
            }

            return computed;
        }

        private void InvalidateContextSpellingCache(string word)
        {
            string normalized = TextHelper.NormalizeWord(word);
            if (string.IsNullOrWhiteSpace(normalized)) return;

            lock (_ctxSpellingCacheLock)
            {
                _ctxSpellingCache.Remove(normalized);
            }
        }

        private void App_DocumentOpen(Word.Document doc)
        {
            try
            {
                ClearSpellHighlightsIfFlagged(doc, "open");
                TryMigrateLegacySpellHighlights(doc);
            }
            catch (Exception ex)
            {
                Logger.Error("DocumentOpen highlight cleanup error", ex);
            }
        }

        private void App_DocumentBeforeSave(Word.Document doc, ref bool saveAsUI, ref bool cancel)
        {
            try
            {
                ClearSpellHighlightsIfFlagged(doc, "save");
            }
            catch (Exception ex)
            {
                Logger.Error("DocumentBeforeSave highlight cleanup error", ex);
            }
        }

        private void App_DocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            try
            {
                ClearSpellHighlightsIfFlagged(doc, "close");
            }
            catch (Exception ex)
            {
                Logger.Error("DocumentBeforeClose highlight cleanup error", ex);
            }
        }

        // ---------------------------------------------------------------------
        //  CommandBar Injection
        // ---------------------------------------------------------------------

        /// <summary>Insert spelling suggestions directly into the main right-click menu natively.</summary>
        private void InsertSpellingSuggestions(string addToDictLabel, List<string> suggestions)
        {
            try
            {
                var bar = this.Application.CommandBars["Text"];
                if (bar == null) return;

                int insertIndex = 1;

                // Add suggestion buttons directly to the root menu (Native Word style)
                if (suggestions != null && suggestions.Count > 0)
                {
                    bool isFirst = true;
                    foreach (var sug in suggestions)
                    {
                        if (string.IsNullOrWhiteSpace(sug)) continue;

                        var btn = (Office.CommandBarButton)bar.Controls.Add(
                            Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, insertIndex++, true);
                        
                        btn.Caption = sug;
                        btn.Tag = CTX_TAG;
                        
                        // Icon removed for cleaner, native Word look
                        
                        if (isFirst) 
                        {
                            btn.BeginGroup = true; // Separator above the first suggestion
                            isFirst = false;
                        }
                        
                        btn.Visible = true;
                        btn.Click += OnSuggestionClicked;
                        _ctxControls.Add(btn);
                    }
                }
                else
                {
                    // If no suggestions, add a disabled "No suggestions" label
                    var btnNoSug = (Office.CommandBarButton)bar.Controls.Add(
                        Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, insertIndex++, true);
                    btnNoSug.Caption = "(Тавсиялар йўқ)";
                    btnNoSug.Tag = CTX_TAG;
                    btnNoSug.Enabled = false;
                    btnNoSug.BeginGroup = true;
                    btnNoSug.Visible = true;
                    _ctxControls.Add(btnNoSug);
                }

                // Add "Луғатга қўшиш" button immediately below the suggestions
                var addBtn = (Office.CommandBarButton)bar.Controls.Add(
                    Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, insertIndex++, true);
                
                addBtn.Caption = addToDictLabel;
                addBtn.Tag = CTX_TAG;
                addBtn.FaceId = 2948; // Dictionary book icon
                addBtn.BeginGroup = true; // Separator before dictionary button
                addBtn.Visible = true;
                addBtn.Click += OnAddToDictClicked;
                _ctxControls.Add(addBtn);

                // Add "Вариантлар" dialog launcher button below AddToDict
                var optionsBtn = (Office.CommandBarButton)bar.Controls.Add(
                    Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, insertIndex, true);
                
                optionsBtn.Caption = "Вариантлар...";
                optionsBtn.Tag = CTX_TAG + "_VARIANTLAR";
                optionsBtn.FaceId = 793; // AutoCorrect icon (same as Ribbon)
                optionsBtn.BeginGroup = true; // Separator
                optionsBtn.Visible = true;
                optionsBtn.Click += OnOpenSuggestionsDialogClicked;
                _ctxControls.Add(optionsBtn);
            }
            catch (Exception ex)
            {
                Logger.Error("InsertSpellingSuggestions error", ex);
            }
        }

        /// <summary>Insert standalone "Луғатга қўшиш" button (word is correctly spelled).</summary>
        private void InsertAddToDictButton(string label)
        {
            try
            {
                var bar = this.Application.CommandBars["Text"];
                if (bar == null) return;

                int insertIndex = 1;

                var btn = (Office.CommandBarButton)bar.Controls.Add(
                    Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, insertIndex++, true);
                btn.Caption = label;
                btn.Tag = CTX_TAG;
                btn.FaceId = 2948;
                btn.BeginGroup = true;
                btn.Visible = true;
                btn.Click += OnAddToDictClicked;
                _ctxControls.Add(btn);

                var optionsBtn = (Office.CommandBarButton)bar.Controls.Add(
                    Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, insertIndex, true);
                
                optionsBtn.Caption = "Вариантлар...";
                optionsBtn.Tag = CTX_TAG + "_VARIANTLAR";
                optionsBtn.FaceId = 793; // AutoCorrect icon (same as Ribbon)
                optionsBtn.BeginGroup = true; // Separator
                optionsBtn.Visible = true;
                optionsBtn.Click += OnOpenSuggestionsDialogClicked;
                _ctxControls.Add(optionsBtn);
            }
            catch (Exception ex)
            {
                Logger.Error("InsertAddToDictButton error", ex);
            }
        }

        /// <summary>Insert grammar error suggestions into the context menu.</summary>
        private void InsertGrammarSuggestions(List<Models.ErrorEntry> grammarErrors, int sentenceStart)
        {
            try
            {
                var bar = this.Application.CommandBars["Text"];
                if (bar == null) return;

                // Calculate insert index after all existing custom controls
                int insertIndex = _ctxControls.Count + 1;

                // Add "Грамматика:" separator label
                var lblGrammar = (Office.CommandBarButton)bar.Controls.Add(
                    Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, insertIndex++, true);
                lblGrammar.Caption = "Грамматика:";
                lblGrammar.Tag = CTX_TAG;
                lblGrammar.Enabled = false;
                lblGrammar.BeginGroup = true;
                lblGrammar.Visible = true;
                _ctxControls.Add(lblGrammar);

                foreach (var error in grammarErrors)
                {
                    string caption = error.Message ?? error.Context ?? error.Word;
                    if (string.IsNullOrWhiteSpace(caption)) continue;

                    // Truncate long messages
                    if (caption.Length > 60)
                        caption = caption.Substring(0, 57) + "...";

                    if (error.BestSuggestion != null)
                    {
                        // Actionable: show suggestion
                        var sugBtn = (Office.CommandBarButton)bar.Controls.Add(
                            Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, insertIndex++, true);
                        sugBtn.Caption = $"→ {error.BestSuggestion}";
                        int absStart = sentenceStart + Math.Max(0, error.StartIndex);
                        int absEnd = sentenceStart + Math.Max(error.EndIndex, error.StartIndex + 1);
                        sugBtn.Tag = BuildContextTag(absStart, absEnd);
                        sugBtn.FaceId = 548; // green checkmark
                        sugBtn.Visible = true;
                        sugBtn.Click += OnSuggestionClicked;
                        _ctxControls.Add(sugBtn);
                    }
                    else
                    {
                        // Info only: show the error message
                        var infoBtn = (Office.CommandBarButton)bar.Controls.Add(
                            Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, insertIndex++, true);
                        infoBtn.Caption = caption;
                        infoBtn.Tag = CTX_TAG;
                        infoBtn.Enabled = false;
                        infoBtn.Visible = true;
                        _ctxControls.Add(infoBtn);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("InsertGrammarSuggestions error", ex);
            }
        }

        // ---------------------------------------------------------------------
        //  Click Handlers
        // ---------------------------------------------------------------------

        private void OnOpenSuggestionsDialogClicked(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            try
            {
                if (!DocumentHelper.IsDocumentOpen()) return;

                System.Drawing.Image icon = null;
                try { icon = Globals.Ribbons?.UzbekOrfoRibbon?.btnSuggestions?.Image; } catch { }

                var workflow = new Services.SuggestionsWorkflowService(
                    SpellingEngine,
                    GrammarEngine,
                    this.Application,
                    word => AddWordToUserDictionary(word),
                    (fromWord, toWord) => DocumentHelper.ReplaceAllInDocument(fromWord, toWord));

                workflow.Execute(icon);
            }
            catch (Exception ex)
            {
                Logger.Error("Open Suggestions Dialog error", ex);
            }
        }

        private void OnSuggestionClicked(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            try
            {
                string suggestion = NormalizeContextMenuSuggestion(ctrl?.Caption);
                if (string.IsNullOrWhiteSpace(suggestion)) return;

                var range = GetTaggedSuggestionRange(ctrl) ?? GetContextMenuTargetRange();
                if (range == null) return;

                string oldWord = TextHelper.NormalizeWord(range.Text ?? _ctxTargetWord ?? string.Empty);

                DocumentHelper.BeginUndoRecord("Сўзни тузатиш");
                try { DocumentHelper.ReplaceRangeText(range, suggestion); }
                finally { DocumentHelper.EndUndoRecord(); }

                TryClearSelectionUnderline();
                if (!string.IsNullOrWhiteSpace(oldWord))
                    ResolveErrorsForWord(oldWord);
            }
            catch (Exception ex)
            {
                Logger.Error("Suggestion click error", ex);
            }
        }

        private static string NormalizeContextMenuSuggestion(string caption)
        {
            if (string.IsNullOrWhiteSpace(caption)) return null;
            string value = caption.Trim();

            if (value.StartsWith("→", StringComparison.Ordinal))
                value = value.Substring(1).TrimStart();
            else if (value.StartsWith("->", StringComparison.Ordinal))
                value = value.Substring(2).TrimStart();

            return value;
        }

        private static string BuildContextTag(int start, int end)
        {
            if (start < 0 || end <= start) return CTX_TAG;
            return $"{CTX_TAG}|{start}|{end}";
        }

        private Word.Range GetTaggedSuggestionRange(Office.CommandBarButton ctrl)
        {
            try
            {
                string tag = ctrl?.Tag ?? string.Empty;
                if (!tag.StartsWith(CTX_TAG, StringComparison.Ordinal))
                    return null;

                string[] parts = tag.Split('|');
                if (parts.Length < 3)
                    return null;

                if (!int.TryParse(parts[1], out int start) || !int.TryParse(parts[2], out int end))
                    return null;

                var doc = this.Application?.ActiveDocument;
                if (doc == null || start < 0 || end <= start || end > doc.Content.End)
                    return null;

                string activeKey = GetDocumentKey(doc);
                if (!string.IsNullOrWhiteSpace(_ctxTargetDocKey) &&
                    !string.Equals(activeKey, _ctxTargetDocKey, StringComparison.OrdinalIgnoreCase))
                    return null;

                return doc.Range(start, end);
            }
            catch
            {
                return null;
            }
        }

        private void OnAddToDictClicked(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            try
            {
                string word = DocumentHelper.GetSelectedWord();
                if (string.IsNullOrWhiteSpace(word))
                    word = _ctxTargetWord;
                if (string.IsNullOrWhiteSpace(word)) return;

                AddWordToUserDictionary(word, clearSelectionUnderline: true);
            }
            catch (Exception ex)
            {
                Logger.Error("AddToDict click error", ex);
            }
        }

        private void CaptureContextMenuTarget(Word.Selection sel, out string rawWord, out string normalizedWord)
        {
            rawWord = null;
            normalizedWord = null;
            ResetContextMenuTarget();

            try
            {
                Word.Range baseRange = null;
                if (sel != null) baseRange = sel.Range;
                if (baseRange == null) baseRange = this.Application?.Selection?.Range;
                if (baseRange == null) return;

                Word.Range wordRange = null;
                try { wordRange = baseRange.Words[1]; }
                catch { wordRange = baseRange; }
                if (wordRange == null) return;

                rawWord = (wordRange.Text ?? string.Empty).Trim();
                normalizedWord = TextHelper.NormalizeWord(rawWord);
                if (string.IsNullOrWhiteSpace(normalizedWord) || normalizedWord.Length < 2) return;

                _ctxTargetWord = normalizedWord;
                _ctxTargetWordStart = wordRange.Start;
                _ctxTargetWordEnd = wordRange.End;
                _ctxTargetDocKey = GetDocumentKey(wordRange.Document);
            }
            catch (Exception ex) { Logger.Warn($"CaptureContextMenuTarget error: {ex.Message}"); }
        }



        private Word.Range GetContextMenuTargetRange()
        {
            try
            {
                var doc = this.Application?.ActiveDocument;
                if (doc != null && !string.IsNullOrWhiteSpace(_ctxTargetDocKey))
                {
                    string activeKey = GetDocumentKey(doc);
                    if (!string.Equals(activeKey, _ctxTargetDocKey, StringComparison.OrdinalIgnoreCase))
                        return null; // document changed after right-click: abort safely

                    if (_ctxTargetWordStart >= 0 &&
                        _ctxTargetWordEnd > _ctxTargetWordStart &&
                        _ctxTargetWordEnd <= doc.Content.End)
                    {
                        return doc.Range(_ctxTargetWordStart, _ctxTargetWordEnd);
                    }
                }
            }
            catch { }

            try
            {
                return this.Application?.Selection?.Range?.Words[1];
            }
            catch
            {
                return null;
            }
        }

        private static string GetDocumentKey(Word.Document doc)
        {
            if (doc == null) return null;
            try
            {
                if (!string.IsNullOrWhiteSpace(doc.FullName))
                    return doc.FullName;
            }
            catch { }

            try { return doc.Name; }
            catch { return null; }
        }

        private void ResetContextMenuTarget()
        {
            _ctxTargetDocKey = null;
            _ctxTargetWordStart = -1;
            _ctxTargetWordEnd = -1;
            _ctxTargetWord = null;
        }



        public static Models.AddWordResult AddWordToUserDictionary(string word, bool clearSelectionUnderline = true)
        {
            var dict = DictionaryService;
            if (dict == null)
            {
                ToastNotification.ShowWarning("Луғат хизмати юкланмаган.");
                return Models.AddWordResult.Invalid;
            }

            string rawWord = word?.Trim();
            if (string.IsNullOrWhiteSpace(rawWord))
            {
                ToastNotification.ShowWarning("Аввал сўзни танланг.");
                return Models.AddWordResult.Invalid;
            }

            string normalizedWord = TextHelper.NormalizeWord(rawWord);
            if (string.IsNullOrWhiteSpace(normalizedWord))
                normalizedWord = rawWord.ToLowerInvariant();

            // Word context menu can trigger duplicate add callbacks in quick succession.
            // Suppress duplicate processing so the second callback doesn't overwrite "added"
            // with "already exists" for the same click action.
            if (TryGetRecentAddResult(normalizedWord, out var recentResult))
            {
                if (clearSelectionUnderline)
                    TryClearSelectionUnderline();
                ResolveErrorsForWord(rawWord);
                return recentResult;
            }

            var result = dict.AddWord(rawWord);
            Globals.ThisAddIn?.InvalidateContextSpellingCache(normalizedWord);
            RememberAddResult(normalizedWord, result);
            switch (result)
            {
                case Models.AddWordResult.Added:
                    if (clearSelectionUnderline)
                        TryClearSelectionUnderline();
                    ResolveErrorsForWord(rawWord);
                    ToastNotification.Success("Луғатга қўшилди", rawWord);
                    break;

                case Models.AddWordResult.AlreadyInMainDictionary:
                    if (clearSelectionUnderline)
                        TryClearSelectionUnderline();
                    ToastNotification.ShowInfo("Луғатда мавжуд", $"\"{rawWord}\" асосий луғатда бор.");
                    ResolveErrorsForWord(rawWord);
                    break;

                case Models.AddWordResult.AlreadyInUserDictionary:
                    if (clearSelectionUnderline)
                        TryClearSelectionUnderline();
                    ToastNotification.ShowInfo("Луғатда мавжуд", $"\"{rawWord}\" аллақачон қўшилган.");
                    ResolveErrorsForWord(rawWord);
                    break;

                default:
                    ToastNotification.ShowWarning("Сўз қўшилмади", $"\"{rawWord}\" яроқсиз.");
                    break;
            }

            return result;
        }

        private static bool TryGetRecentAddResult(string normalizedWord, out Models.AddWordResult result)
        {
            lock (_addWordRequestLock)
            {
                bool isDuplicate =
                    !string.IsNullOrWhiteSpace(_lastAddRequestWord) &&
                    _lastAddRequestAtUtc != DateTime.MinValue &&
                    string.Equals(_lastAddRequestWord, normalizedWord, StringComparison.OrdinalIgnoreCase) &&
                    (DateTime.UtcNow - _lastAddRequestAtUtc) <= _duplicateAddRequestWindow;

                if (isDuplicate)
                {
                    result = _lastAddRequestResult;
                    return true;
                }
            }

            result = Models.AddWordResult.Invalid;
            return false;
        }

        private static void RememberAddResult(string normalizedWord, Models.AddWordResult result)
        {
            lock (_addWordRequestLock)
            {
                _lastAddRequestWord = normalizedWord;
                _lastAddRequestAtUtc = DateTime.UtcNow;
                _lastAddRequestResult = result;
            }
        }

        private static void TryClearSelectionUnderline()
        {
            try
            {
                var sel = Globals.ThisAddIn?.Application?.Selection;
                if (sel?.Range == null) return;
                sel.Range.Underline = Word.WdUnderline.wdUnderlineNone;
                sel.Range.Font.UnderlineColor = Word.WdColor.wdColorAutomatic;
            }
            catch { }
        }

        private static void ResolveErrorsForWord(string word)
        {
            var store = ErrorStore;
            if (store == null || string.IsNullOrWhiteSpace(word)) return;

            string normalizedTarget = TextHelper.NormalizeWord(word);
            var matching = store.Errors
                .Where(err =>
                    err != null &&
                    !string.IsNullOrWhiteSpace(err.Word) &&
                    (
                        err.Word.Trim().Equals(word, StringComparison.OrdinalIgnoreCase) ||
                        TextHelper.NormalizeWord(err.Word).Equals(normalizedTarget, StringComparison.OrdinalIgnoreCase)
                    ))
                .ToList();

            foreach (var m in matching)
            {
                try
                {
                    if (m.Range != null)
                    {
                        m.Range.Underline = Word.WdUnderline.wdUnderlineNone;
                        m.Range.Font.UnderlineColor = Word.WdColor.wdColorAutomatic;
                    }
                }
                catch { }

                m.IsResolved = true;
                store.RemoveError(m);
            }
        }

        public static void SetSpellHighlightFlag(Word.Document doc, bool hasHighlights)
        {
            if (doc == null) return;

            try
            {
                Word.Variable flagVar = null;
                try { flagVar = doc.Variables[SPELL_HIGHLIGHT_FLAG_VAR]; }
                catch { }

                if (flagVar == null)
                {
                    if (hasHighlights)
                        doc.Variables.Add(SPELL_HIGHLIGHT_FLAG_VAR, "1");
                }
                else
                {
                    string desired = hasHighlights ? "1" : "0";
                    string current = null;
                    try { current = flagVar.Value; } catch { }
                    if (!string.Equals((current ?? string.Empty).Trim(), desired, StringComparison.Ordinal))
                        flagVar.Value = desired;
                }
            }
            catch { }
        }

        private static bool HasSpellHighlightFlag(Word.Document doc)
        {
            if (doc == null) return false;

            try
            {
                string value = null;
                try { value = doc.Variables[SPELL_HIGHLIGHT_FLAG_VAR]?.Value; }
                catch { }

                if (string.IsNullOrWhiteSpace(value)) return false;
                value = value.Trim();
                return value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static void ClearSpellHighlightsIfFlagged(Word.Document doc, string reason)
        {
            if (doc == null) return;
            if (!HasSpellHighlightFlag(doc)) return;

            try
            {
                int clearedFromStore = ClearHighlightsFromStoreRanges(doc);
                if (clearedFromStore == 0)
                    SpellingEngine?.ClearHighlights(doc);

                RemoveErrorsForDocument(doc);
                // Avoid modifying freshly opened documents just to flip an internal flag.
                // We persist flag reset only during close/save flows.
                if (!string.Equals(reason, "open", StringComparison.OrdinalIgnoreCase))
                    SetSpellHighlightFlag(doc, false);
                Logger.Info($"Spell highlights cleaned on {reason}: {doc.Name}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Spell highlight cleanup failed on {reason}", ex);
            }
        }

        private static void TryMigrateLegacySpellHighlights(Word.Document doc)
        {
            if (doc == null) return;
            if (HasSpellHighlightFlag(doc)) return;
            if (HasLegacyHighlightMigrationFlag(doc)) return;

            try
            {
                if (!ContainsWavyUnderlineFormatting(doc))
                    return;

                SpellingEngine?.ClearHighlights(doc);
                RemoveErrorsForDocument(doc);
                SetSpellHighlightFlag(doc, false);
                SetLegacyHighlightMigrationFlag(doc, true);
                Logger.Info($"Legacy spell highlights migrated: {doc.Name}");
            }
            catch (Exception ex)
            {
                Logger.Error("Legacy spell highlight migration failed", ex);
            }
        }

        private static bool ContainsWavyUnderlineFormatting(Word.Document doc)
        {
            try
            {
                if (doc?.Content == null) return false;

                var range = doc.Content.Duplicate;
                var find = range.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = string.Empty;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.MatchWildcards = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchAllWordForms = false;
                find.MatchSoundsLike = false;
                find.Format = true;
                find.Font.Underline = Word.WdUnderline.wdUnderlineWavy;

                return find.Execute();
            }
            catch
            {
                return false;
            }
        }

        private static bool HasLegacyHighlightMigrationFlag(Word.Document doc)
        {
            if (doc == null) return false;
            try
            {
                string value = null;
                try { value = doc.Variables[SPELL_HIGHLIGHT_MIGRATED_VAR]?.Value; }
                catch { }
                if (string.IsNullOrWhiteSpace(value)) return false;
                value = value.Trim();
                return value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static void SetLegacyHighlightMigrationFlag(Word.Document doc, bool migrated)
        {
            if (doc == null) return;
            try
            {
                Word.Variable flagVar = null;
                try { flagVar = doc.Variables[SPELL_HIGHLIGHT_MIGRATED_VAR]; }
                catch { }

                if (flagVar == null)
                    doc.Variables.Add(SPELL_HIGHLIGHT_MIGRATED_VAR, migrated ? "1" : "0");
                else
                    flagVar.Value = migrated ? "1" : "0";
            }
            catch { }
        }

        private static void RemoveErrorsForDocument(Word.Document doc)
        {
            if (doc == null) return;
            var store = ErrorStore;
            if (store == null) return;

            var snapshot = store.Errors;
            var toRemove = new List<Models.ErrorEntry>();

            foreach (var err in snapshot)
            {
                try
                {
                    if (err?.Range != null && IsSameDocument(err.Range.Document, doc))
                        toRemove.Add(err);
                }
                catch { }
            }

            foreach (var err in toRemove)
            {
                try { store.RemoveError(err); } catch { }
            }
        }

        private static int ClearHighlightsFromStoreRanges(Word.Document doc)
        {
            if (doc == null) return 0;
            var store = ErrorStore;
            if (store == null) return 0;

            int cleared = 0;
            foreach (var err in store.Errors)
            {
                try
                {
                    if (err?.Range == null) continue;
                    if (!IsSameDocument(err.Range.Document, doc)) continue;

                    err.Range.Underline = Word.WdUnderline.wdUnderlineNone;
                    err.Range.Font.UnderlineColor = Word.WdColor.wdColorAutomatic;
                    err.IsResolved = true;
                    cleared++;
                }
                catch { }
            }

            return cleared;
        }

        private static bool IsSameDocument(Word.Document a, Word.Document b)
        {
            if (a == null || b == null) return false;

            try
            {
                string aFull = a.FullName;
                string bFull = b.FullName;
                if (!string.IsNullOrWhiteSpace(aFull) && !string.IsNullOrWhiteSpace(bFull))
                    return string.Equals(aFull, bFull, StringComparison.OrdinalIgnoreCase);
            }
            catch { }

            try { return a == b; }
            catch { return false; }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
