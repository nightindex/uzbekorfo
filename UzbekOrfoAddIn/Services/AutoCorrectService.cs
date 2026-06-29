using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Provides real-time auto-correction while the user types.
    /// Hooks into Word's WindowSelectionChange event to detect when a user
    /// finishes typing a word (moves to next word) and auto-corrects the previous word.
    /// </summary>
    public class AutoCorrectService : IDisposable
    {
        private readonly SpellingEngine _engine;
        private readonly DictionaryService _dictionary;
        private Word.Application _app;
        private bool _isEnabled;
        private bool _isProcessing;
        private string _lastCheckedWordKey;
        private string _lastSelectionSnapshot;
        private readonly Timer _selectionPollTimer;

        // Deferred correction — queued from TryAutoCorrect, executed on Application.Idle
        private Word.Range _pendingRange;
        private string _pendingText;
        private string _pendingLogOriginal;
        private double _pendingLogConfidence;

        /// <summary>Whether real-time auto-correct is currently active.</summary>
        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_isEnabled == value) return;
                _isEnabled = value;
                if (_isEnabled)
                {
                    _lastCheckedWordKey = null;
                    _lastSelectionSnapshot = null;
                    Attach();
                }
                else
                {
                    Detach();
                }
            }
        }

        public AutoCorrectService(SpellingEngine engine, DictionaryService dictionary)
        {
            _engine = engine ?? throw new ArgumentNullException(nameof(engine));
            _dictionary = dictionary ?? throw new ArgumentNullException(nameof(dictionary));

            // Fallback polling makes auto-correct reliable across Word builds where
            // SelectionChange timing around Space/Tab/Enter is inconsistent.
            _selectionPollTimer = new Timer { Interval = 500 };
            _selectionPollTimer.Tick += OnSelectionPollTick;
        }

        /// <summary>
        /// Initializes the service by storing the Word Application reference.
        /// Call this once during add-in startup.
        /// </summary>
        public void Initialize(Word.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        /// <summary>
        /// Attaches to Word's document change event to monitor typing.
        /// </summary>
        private void Attach()
        {
            if (_app == null) return;
            try
            {
                _app.WindowSelectionChange += OnSelectionChange;
                _selectionPollTimer.Start();
                Logger.Info("РђРІС‚Рѕ С‚СѓР·Р°С‚РёС€: WindowSelectionChange С…РѕРґРёСЃР°СЃРёРіР° СѓР»Р°РЅРґРё.");
            }
            catch (Exception ex)
            {
                Logger.Error("РђРІС‚Рѕ С‚СѓР·Р°С‚РёС€: СѓР»Р°РЅРёС€РґР° С…Р°С‚Рѕ", ex);
            }
        }

        /// <summary>
        /// Detaches from Word's document change event.
        /// </summary>
        private void Detach()
        {
            if (_app == null) return;
            try
            {
                _selectionPollTimer.Stop();
                _app.WindowSelectionChange -= OnSelectionChange;
                Logger.Info("РђРІС‚Рѕ С‚СѓР·Р°С‚РёС€: С…РѕРґРёСЃР°РґР°РЅ Р°Р¶СЂР°Р»РґРё.");
            }
            catch (Exception ex)
            {
                Logger.Error("РђРІС‚Рѕ С‚СѓР·Р°С‚РёС€: Р°Р¶СЂР°Р»РёС€РґР° С…Р°С‚Рѕ", ex);
            }
        }

        /// <summary>
        /// Fired every time the selection changes in the active document.
        /// We use this to check the word the cursor just left.
        /// </summary>
        private void OnSelectionChange(Word.Selection sel)
        {
            TryAutoCorrect(sel, null);
        }

        private void OnSelectionPollTick(object sender, EventArgs e)
        {
            if (!_isEnabled || _app == null || _isProcessing) return;

            Word.Selection sel = null;
            try { sel = _app.Selection; } catch { }
            if (sel == null) return;

            string snapshot = BuildSelectionSnapshot(sel);
            if (string.Equals(snapshot, _lastSelectionSnapshot, StringComparison.Ordinal))
                return;

            TryAutoCorrect(sel, snapshot);
        }

        private void TryAutoCorrect(Word.Selection sel, string knownSnapshot)
        {
            if (!_isEnabled || _isProcessing || sel == null) return;

            try
            {
                _isProcessing = true;
                _lastSelectionSnapshot = knownSnapshot ?? BuildSelectionSnapshot(sel);

                // We only auto-correct when the cursor is in an insertion point
                // (not selecting a block of text)
                if (sel.Type != Word.WdSelectionType.wdSelectionIP &&
                    sel.Type != Word.WdSelectionType.wdSelectionNormal)
                    return;

                // Auto-correct only when a word boundary is typed (Space/Enter/Tab/punctuation),
                // not while user is still typing inside a word.
                if (!IsBoundaryContext(sel)) return;

                // Get the word immediately before the current cursor position
                Word.Range prevWordRange = GetPreviousWord(sel);
                if (prevWordRange == null) return;

                string prevWord = prevWordRange.Text;
                if (string.IsNullOrWhiteSpace(prevWord))
                {
                    Marshal.ReleaseComObject(prevWordRange);
                    return;
                }

                string trimmed = prevWord.Trim();

                // Avoid re-checking the same word repeatedly
                string wordKey = $"{prevWordRange.Start}:{prevWordRange.End}:{trimmed}";
                if (string.Equals(wordKey, _lastCheckedWordKey, StringComparison.Ordinal))
                {
                    Marshal.ReleaseComObject(prevWordRange);
                    return;
                }

                _lastCheckedWordKey = wordKey;

                // Skip words that should be ignored (numbers, urls, etc.)
                var normalized = TextHelper.NormalizeWord(trimmed);
                if (TextHelper.ShouldSkipWord(normalized)) { Marshal.ReleaseComObject(prevWordRange); return; }
                if (normalized.Length <= 1) { Marshal.ReleaseComObject(prevWordRange); return; }

                // Check if word is correct
                if (_engine.IsCorrect(trimmed)) { Marshal.ReleaseComObject(prevWordRange); return; }

                // Get the best suggestion
                var suggestions = _engine.GetSuggestions(trimmed, 1);
                if (suggestions == null || suggestions.Count == 0) { Marshal.ReleaseComObject(prevWordRange); return; }

                var best = suggestions.First();
                if (!ShouldAutoCorrect(best.Confidence, best.EditDistance)) { Marshal.ReleaseComObject(prevWordRange); return; }

                // If a correction is already queued but not yet applied, skip this one.
                // Overwriting _pendingRange would silently drop the first correction AND
                // register Application.Idle a second time (double-fire).
                if (_pendingRange != null) { Marshal.ReleaseComObject(prevWordRange); return; }

                // Defer the COM mutation out of the current call frame.
                // Writing prevWordRange.Text synchronously here fires
                // WindowSelectionChange (and other COM events) re-entrantly
                // while the timer/event callback stack frame is still active.
                // Third-party add-in handlers (e.g. Lightkey) invoked inside
                // that corrupted call frame cause a FatalExecutionEngineError.
                // By posting the write to Application.Idle we ensure it runs
                // from a clean top-level message-pump cycle.
                _pendingRange = prevWordRange;
                _pendingText = best.Text;
                _pendingLogOriginal = trimmed;
                _pendingLogConfidence = best.Confidence;
                System.Windows.Forms.Application.Idle += OnApplyDeferredCorrection;
            }
            catch (Exception ex)
            {
                Logger.Error("РђРІС‚Рѕ С‚СѓР·Р°С‚РёС€: SelectionChange РёС€Р»РѕРІС‡РёСЃРёРґР° С…Р°С‚Рѕ", ex);
            }
            finally
            {
                _isProcessing = false;
            }
        }

        /// <summary>
        /// Applies the queued auto-correction from a clean message-pump frame.
        /// Runs exactly once per queued correction.
        /// </summary>
        private void OnApplyDeferredCorrection(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Idle -= OnApplyDeferredCorrection;

            var range = _pendingRange;
            var text = _pendingText;
            var logOriginal = _pendingLogOriginal;
            var logConfidence = _pendingLogConfidence;
            _pendingRange = null;
            _pendingText = null;
            _pendingLogOriginal = null;

            if (range == null || string.IsNullOrEmpty(text)) return;

            try
            {
                DocumentHelper.BeginUndoRecord("Auto-correct");
                try
                {
                    range.Text = text;
                }
                finally
                {
                    DocumentHelper.EndUndoRecord();
                }

                Logger.Info($"Auto-correct: \"{logOriginal}\" -> \"{text}\" (conf: {logConfidence:P0})");
            }
            catch (Exception ex)
            {
                Logger.Error($"Auto-correct: \"{logOriginal}\" replacement error", ex);
            }
            finally
            {
                try { Marshal.ReleaseComObject(range); } catch { }
            }
        }

        private static string BuildSelectionSnapshot(Word.Selection sel)
        {
            try
            {
                if (sel == null || sel.Document == null) return string.Empty;

                var doc = sel.Document;
                string docKey = string.Empty;
                try
                {
                    docKey = doc.FullName;
                    if (string.IsNullOrEmpty(docKey))
                        docKey = doc.Name ?? string.Empty;
                }
                catch { }

                int before1 = -1;
                int before2 = -1;
                if (sel.Start > 0)
                {
                    int left = Math.Max(0, sel.Start - 2);
                    var beforeRange = doc.Range(left, sel.Start);
                    try
                    {
                        string tail = beforeRange.Text ?? string.Empty;
                        if (tail.Length >= 1) before1 = tail[tail.Length - 1];
                        if (tail.Length >= 2) before2 = tail[tail.Length - 2];
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(beforeRange);
                    }
                }

                return $"{docKey}|{sel.Start}|{sel.End}|{before1}|{before2}";
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Gets the word range immediately before the current cursor position.
        /// Returns null if not applicable.
        /// </summary>
        private Word.Range GetPreviousWord(Word.Selection sel)
        {
            try
            {
                // If the selection is at the very start, there's no previous word
                if (sel.Start <= 1) return null;

                var doc = sel.Document;
                if (doc == null) return null;

                int cursorPos = sel.Start;
                int windowStart = Math.Max(0, cursorPos - 128);
                var tailRange = doc.Range(windowStart, cursorPos);
                string tail;
                try
                {
                    tail = tailRange.Text;
                }
                finally
                {
                    Marshal.ReleaseComObject(tailRange);
                }
                if (string.IsNullOrEmpty(tail)) return null;

                // 1) Skip trailing spaces/newlines (cursor after Space/Enter lands here).
                int i = tail.Length - 1;
                while (i >= 0 && char.IsWhiteSpace(tail[i])) i--;
                if (i < 0) return null;

                // 2) Skip punctuation right before cursor (e.g., "word, ").
                while (i >= 0 && !IsWordChar(tail[i])) i--;
                if (i < 0) return null;

                // 3) Capture the word itself.
                int end = i + 1;
                while (i >= 0 && IsWordChar(tail[i])) i--;
                int start = i + 1;
                if (end <= start) return null;

                int absStart = windowStart + start;
                int absEnd = windowStart + end;
                if (absEnd <= absStart) return null;

                // NOTE: wordRange is intentionally NOT released here — it is returned
                // to the caller (TryAutoCorrect) which uses it for replacement and
                // is responsible for releasing it.
                var wordRange = doc.Range(absStart, absEnd);
                string text = wordRange.Text;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    return wordRange;
                }

                Marshal.ReleaseComObject(wordRange);
                return null;
            }
            catch
            {
                return null;
            }
        }

        private static bool IsWordChar(char c)
        {
            return char.IsLetter(c) ||
                   c == '\'' ||
                   c == '\u02BB' || // К»
                   c == '\u02BC' || // Кј
                   c == '\u2018' || // вЂ
                   c == '\u2019';   // вЂ™
        }

        private static bool IsBoundaryContext(Word.Selection sel)
        {
            try
            {
                if (sel == null || sel.Document == null) return false;
                if (sel.Start <= 0) return false;

                var beforeRange = sel.Document.Range(sel.Start - 1, sel.Start);
                try
                {
                    string before = beforeRange.Text;
                    if (string.IsNullOrEmpty(before)) return false;

                    char last = before[before.Length - 1];
                    return char.IsWhiteSpace(last) || !IsWordChar(last);
                }
                finally
                {
                    Marshal.ReleaseComObject(beforeRange);
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool ShouldAutoCorrect(double confidence, int editDistance)
        {
            if (editDistance <= 1)
                return confidence >= 0.52;
            if (editDistance == 2)
                return confidence >= 0.68;
            return confidence >= 0.82;
        }

        public void Dispose()
        {
            Detach();
            // Unhook any pending deferred correction
            try { System.Windows.Forms.Application.Idle -= OnApplyDeferredCorrection; } catch { }
            _pendingRange = null;
            _pendingText = null;
            _pendingLogOriginal = null;
            _selectionPollTimer.Tick -= OnSelectionPollTick;
            _selectionPollTimer.Dispose();
            _app = null;
        }
    }
}
