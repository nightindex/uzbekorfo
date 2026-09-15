using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Helpers;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Services
{
    public sealed class WordCompletionController : IDisposable
    {
        private readonly Word.Application _app;
        private readonly SettingsManager _settings;
        private readonly DictionaryService _dictionary;
        private readonly AutoCorrectService _autoCorrect;
        private readonly UzbekMorphAnalyzer _morphology;
        private readonly CompletionPreferences _preferences;
        private readonly System.Windows.Forms.Timer _timer = new System.Windows.Forms.Timer { Interval = 40 };
        private readonly CompletionPopup _popup = new CompletionPopup();
        private readonly GhostSuggestionWindow _ghost = new GhostSuggestionWindow();
        private readonly Control _dispatcher = new Control();
        private Task<WordCompletionEngine> _build;
        private Task<string[]> _query;
        private WordCompletionEngine _engine;
        private WordCompletionContext _current;
        private WordCompletionContext _queryContext;
        private CancellationTokenSource _cancel;
        private CancellationTokenSource _buildCancel;
        private bool _disposed;
        private string[] _shown = new string[0];
        public bool PopupVisible => !_disposed && _popup.Visible;
        public bool GhostVisible => !_disposed && _ghost.Visible;
        public bool SuggestionVisible => PopupVisible || GhostVisible;
        public bool CanHandleKeys => SuggestionVisible && HasCurrentEditingFocus();
        public bool CanAcceptSuggestion => SuggestionVisible && HasCurrentEditingFocus();
        public bool CanOpenAlternatives => SuggestionVisible && _shown.Length > 0 && HasCurrentEditingFocus();
        public bool CanNavigateAlternatives => PopupVisible && HasCurrentEditingFocus();
        public bool IsReady => _engine != null;
        public bool IsEnabled => _settings.PredictionsEnabled;

        public WordCompletionController(Word.Application app, SettingsManager settings, DictionaryService dictionary,
            AutoCorrectService autoCorrect, UzbekMorphAnalyzer morphology)
        {
            _app = app; _settings = settings; _dictionary = dictionary; _autoCorrect = autoCorrect; _morphology = morphology;
            _preferences = new CompletionPreferences(settings.MatnAiPreferencesPath);
            if (settings.MatnAiLearningConsent)
                try { _preferences.Load(); } catch { Logger.Warn("MatnAI шахсий созламалари очилмади; стандарт тартиблаш қўлланади."); }
            _timer.Tick += Tick;
            _popup.Accepted += Accept;
            _popup.Dismissed += Dismiss;
            var dispatcherHandle = _dispatcher.Handle;
            _dictionary.VocabularySaved += OnVocabularySaved;
            if (IsEnabled) { Rebuild(); _timer.Start(); }
        }

        public void SetEnabled(bool enabled)
        {
            _settings.PredictionsEnabled = enabled;
            _settings.Save();
            Clear();
            if (enabled) { Rebuild(); _timer.Start(); }
            else _timer.Stop();
        }

        public void SetLearningEnabled(bool enabled)
        {
            _settings.MatnAiLearningConsent = enabled;
            if (enabled)
                try { _preferences.Load(); } catch { Logger.Warn("MatnAI шахсий созламаларини юклаб бўлмади."); }
            _settings.Save();
            Dismiss();
        }

        public void ResetLearning()
        {
            _preferences.Clear();
            Dismiss();
        }

        public void Rebuild()
        {
            if (_disposed) return;
            _buildCancel?.Cancel();
            _buildCancel?.Dispose();
            _buildCancel = new CancellationTokenSource();
            var cancellation = _buildCancel.Token;
            // Capture on the owning UI thread. No worker enumerates the mutable dictionary.
            var words = _dictionary.GetCompletionSnapshot();
            _engine = null;
            Clear();
            string rulesPath = _settings.SuffixesPath;
            _build = Task.Run(() =>
            {
                cancellation.ThrowIfCancellationRequested();
                string[] endings;
                try
                {
                    endings = MorphologyRuleSetLoader.Load(rulesPath).Suffixes.Where(s => s.Productive)
                        .SelectMany(s => new[] { s.Latin, s.Cyrillic }).ToArray();
                }
                catch { endings = new string[0]; } // Invalid rules: lexical completion only.
                return new WordCompletionEngine(words, endings, cancellation: cancellation);
            }, cancellation);
            if (IsEnabled) _timer.Start();
        }

        private void OnVocabularySaved()
        {
            if (_disposed || (_engine == null && _build == null)) return;
            // Dictionary import may finish on another thread. Capture/rebuild only on Word's STA.
            try { _dispatcher.BeginInvoke((Action)(() => { if (!_disposed) Rebuild(); })); }
            catch (InvalidOperationException) { /* Word is closing. */ }
        }

        public void Dismiss()
        {
            _popup.Hide();
            _ghost.Hide();
            _shown = new string[0];
            _cancel?.Cancel();
            // Keep current identity, suppress redisplay until the user edits/moves.
            if (!IsEnabled) _timer.Stop();
        }

        private void Clear()
        {
            _popup.Hide();
            _ghost.Hide();
            _shown = new string[0];
            _cancel?.Cancel();
            _current?.Dispose(); _current = null;
        }

        private void Tick(object sender, EventArgs args)
        {
            if (_disposed) return;
            try
            {
                if (!IsEnabled && _query == null && !SuggestionVisible) { _timer.Stop(); return; }
                if (_build != null && _build.IsCompleted)
                {
                    var completed = _build; _build = null;
                    if (completed.Status == TaskStatus.RanToCompletion) _engine = completed.Result;
                    else { Logger.Warn("MatnAI индекси тайёр эмас."); _timer.Stop(); return; }
                }
                if (!WordCompletionContext.HasEditingFocus(_app)) { Clear(); return; }
                using (var latest = WordCompletionContext.Capture(_app))
                {
                    if (latest == null) { Clear(); return; }
                    if (_current == null || !_current.SameAs(latest))
                    {
                        Clear();
                        _current = WordCompletionContext.Capture(_app);
                        if (_current == null) return;
                    }
                }
                if (SuggestionVisible && _current != null)
                    PositionPresentation();
                if (_query != null)
                {
                    if (!_query.IsCompleted) return; // At most one active query.
                    var finished = _query; _query = null;
                    bool valid = !_cancel.IsCancellationRequested && _current != null && _current.SameAs(_queryContext);
                    _queryContext.Dispose(); _queryContext = null;
                    _cancel.Dispose(); _cancel = null;
                    if (valid && finished.Status == TaskStatus.RanToCompletion && finished.Result.Length > 0)
                    {
                        // Worker generates bounded proposals only. The SAME runtime morphology
                        // engine validates them here, on the dictionary's owning UI thread.
                        _shown = WordCompletionEngine.ValidateCandidates(finished.Result,
                            candidate => _dictionary.Contains(candidate) || _morphology.Analyze(candidate).IsValidInflectedForm,
                            _settings.MaxPredictions);
                        if (_shown.Length > 0) PresentPrimarySuggestion();
                        else Dismiss();
                    }
                    else if (finished.IsFaulted) Logger.Warn("MatnAI таклифи тайёрланмади.");
                    return;
                }
                if (_engine == null || _current == null ||
                    _current.Prefix.Count(char.IsLetter) < _settings.MinPredictionLength ||
                    _requested == _current) return;
                if (!IsEnabled) { _timer.Stop(); return; }
                _requested = _current;
                _queryContext = WordCompletionContext.Capture(_app);
                if (_queryContext == null) return;
                string prefix = _queryContext.Prefix;
                var engine = _engine;
                var preferences = _settings.MatnAiLearningConsent ? _preferences.Snapshot() : null;
                _cancel = new CancellationTokenSource();
                var token = _cancel.Token;
                _query = Task.Run(() => engine.Complete(prefix, 10, token, deferMorphologyValidation: true,
                    acceptanceCounts: preferences), token);
            }
            catch (Exception ex)
            {
                Clear();
                Logger.Warn("MatnAI хост хатосидан кейин тўхтатилди: " + ex.GetType().Name);
                _timer.Stop(); // Avoid a repeated COM failure/log loop.
            }
        }
        private WordCompletionContext _requested;

        private bool HasCurrentEditingFocus()
        {
            return _current != null && WordCompletionContext.HasEditingFocus(_current.WindowHandle);
        }

        private void PresentPrimarySuggestion()
        {
            _popup.Hide();
            string tail = _shown.Length == 0 || _current == null
                ? null : WordCompletionEngine.GetGhostTail(
                    _current.Prefix, _shown[0], _settings.MinPredictionLength);
            if (tail == null)
            {
                _ghost.Hide();
                return;
            }
            try
            {
                var anchor = _current.GetCaretAnchor(_app);
                _ghost.Present(tail, anchor, new WindowOwner(_current.WindowHandle));
            }
            catch (System.Runtime.InteropServices.COMException) { _ghost.Hide(); }
        }

        private void PositionPresentation()
        {
            try
            {
                var anchor = _current.GetCaretAnchor(_app);
                var owner = new WindowOwner(_current.WindowHandle);
                if (_popup.Visible) _popup.Present(_shown, anchor, owner);
                else if (_ghost.Visible)
                {
                    string tail = _shown.Length == 0 ? null :
                        WordCompletionEngine.GetGhostTail(
                            _current.Prefix, _shown[0], _settings.MinPredictionLength);
                    if (tail == null || !_ghost.Present(tail, anchor, owner)) _ghost.Hide();
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                _popup.Hide();
                _ghost.Hide();
            }
        }

        public void ShowAlternatives()
        {
            if (!CanOpenAlternatives) return;
            try
            {
                _ghost.Hide();
                var anchor = _current.GetCaretAnchor(_app);
                _popup.Present(_shown, anchor, new WindowOwner(_current.WindowHandle));
            }
            catch (System.Runtime.InteropServices.COMException) { _popup.Hide(); }
        }

        public void AcceptSelected()
        {
            if (!CanAcceptSuggestion) return;
            Accept(_popup.Visible ? _popup.Selected : _shown.FirstOrDefault());
        }

        public void MoveSelection(int direction)
        {
            if (!_popup.Visible)
            {
                ShowAlternatives();
                return;
            }
            if (CanNavigateAlternatives) _popup.MoveSelection(direction);
        }

        private void Accept(string word)
        {
            if (_current == null || !SuggestionVisible || string.IsNullOrEmpty(word)) return;
            bool autoCorrect = _autoCorrect.IsEnabled;
            try
            {
                _autoCorrect.IsEnabled = false; // Clears pending correction and prevents reentrant edits.
                if (_current.TryInsert(_app, word) && _settings.MatnAiLearningConsent)
                {
                    _preferences.Record(word);
                    // Persist on explicit acceptance only, never on ordinary keystrokes.
                    try { _preferences.Save(); } catch { Logger.Warn("MatnAI шахсий созламаларини сақлаб бўлмади."); }
                }
            }
            catch (Exception ex) { Logger.Warn("MatnAI таклифини қўшиб бўлмади: " + ex.GetType().Name); }
            finally { _autoCorrect.IsEnabled = autoCorrect; Dismiss(); }
        }
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _dictionary.VocabularySaved -= OnVocabularySaved;
            _timer.Stop(); _timer.Dispose(); Clear();
            _queryContext?.Dispose(); _queryContext = null;
            _buildCancel?.Cancel(); _buildCancel?.Dispose();
            _cancel?.Dispose();
            _popup.Dispose();
            _ghost.Dispose();
            _dispatcher.Dispose();
        }
        private sealed class WindowOwner : IWin32Window
        {
            public IntPtr Handle { get; }
            public WindowOwner(int hwnd) { Handle = new IntPtr(hwnd); }
        }
    }
}
