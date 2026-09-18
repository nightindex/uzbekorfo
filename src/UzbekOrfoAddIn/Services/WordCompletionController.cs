using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Prediction;
using UzbekOrfoAddIn.Models;
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
        private readonly AcceptanceStore _preferences;
        private readonly PredictionMetrics _metrics;
        private readonly Stopwatch _latency = new Stopwatch();
        private long _lastAcceptTimestamp;
        private bool _undoPending;
        private WordCompletionContext _beforeAcceptance;
        public CollectionStore Collections { get; }
        private readonly Dictionary<long, string[]> _documentCollections = new Dictionary<long, string[]>();
        private Task<CollectionSnapshot> _collectionBuild;
        private CollectionSnapshot _collections;
        private int _collectionRevision;
        private int _queryRevision;
        private long _nextRequest;
        private string _displayedCompletion;
        private Dictionary<string, string> _provenance = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly System.Windows.Forms.Timer _timer = new System.Windows.Forms.Timer { Interval = 40 };
        private readonly CompletionPopup _popup = new CompletionPopup();
        private readonly GhostSuggestionWindow _ghost = new GhostSuggestionWindow();
        private readonly Control _dispatcher = new Control();
        private Task<WordCompletionEngine> _build;
        private Task<QueryBatch> _query;
        private WordCompletionEngine _engine;
        private WordCompletionContext _current;
        private WordCompletionContext _queryContext;
        private CancellationTokenSource _cancel;
        private CancellationTokenSource _buildCancel;
        private bool _disposed;
        private string[] _shown = new string[0];
        private string _continuedWord;
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
            AutoCorrectService autoCorrect, UzbekMorphAnalyzer morphology, CollectionStore collections = null)
        {
            _app = app; _settings = settings; _dictionary = dictionary; _autoCorrect = autoCorrect; _morphology = morphology;
            Collections = collections ?? new CollectionStore();
            _metrics = new PredictionMetrics(Collections.RootDirectory) { Enabled = settings.MatnAiMetricsConsent };
            HotkeyManager.KeyObserved += OnKeyObserved;
            _preferences = new AcceptanceStore(Collections.RootDirectory);
            if (settings.MatnAiLearningConsent)
                try { _preferences.Load(settings.MatnAiPreferencesPath); } catch { Logger.Warn("MatnAI шахсий созламалари очилмади; стандарт тартиблаш қўлланади."); }
            _timer.Tick += Tick;
            _popup.Accepted += Accept;
            _popup.Dismissed += Dismiss;
            var dispatcherHandle = _dispatcher.Handle;
            _dictionary.VocabularySaved += OnVocabularySaved;
            ReloadCollections();
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
                try { _preferences.Load(_settings.MatnAiPreferencesPath); } catch { Logger.Warn("MatnAI шахсий созламаларини юклаб бўлмади."); }
            _settings.Save();
            HideSuggestion();
        }

        public void ResetLearning()
        {
            _preferences.Clear(_settings.MatnAiPreferencesPath);
            HideSuggestion();
        }

        public void SetMetricsEnabled(bool enabled)
        {
            _settings.MatnAiMetricsConsent = enabled;
            _settings.Save(); _metrics.Enabled = enabled;
        }

        private void OnKeyObserved(Keys key)
        {
            if (!_settings.MatnAiMetricsConsent || _beforeAcceptance == null ||
                !WordCompletionContext.HasEditingFocus(_beforeAcceptance.WindowHandle)) return;
            if (key == (Keys.Control | Keys.Z) && Stopwatch.GetTimestamp() - _lastAcceptTimestamp < Stopwatch.Frequency * 5)
                _undoPending = true;
            else { _lastAcceptTimestamp = 0; }
        }

        private static long DocumentKey(Word.Document document)
        {
            IntPtr identity = Marshal.GetIUnknownForObject(document);
            try { return identity.ToInt64(); }
            finally { Marshal.Release(identity); }
        }

        public string[] GetActiveCollectionIds()
        {
            try
            {
                string[] ids;
                return _documentCollections.TryGetValue(DocumentKey(_app.ActiveDocument), out ids) ? ids.ToArray() : new string[0];
            }
            catch (COMException) { return new string[0]; }
        }

        public void SetActiveCollectionIds(string[] ids)
        {
            try { _documentCollections[DocumentKey(_app.ActiveDocument)] = (ids ?? new string[0]).Distinct().ToArray(); }
            catch (COMException) { return; }
            _collectionRevision++;
            Clear();
        }

        public void ForgetDocument(Word.Document document)
        {
            _documentCollections.Remove(DocumentKey(document));
            Clear();
        }
        public void ForgetCollectionLearning(string id)
        {
            if (!_settings.MatnAiLearningConsent) _preferences.Load(_settings.MatnAiPreferencesPath);
            _preferences.RemoveCollection(id);
        }

        public void ReloadCollections()
        {
            _collectionRevision++;
            Clear();
            _collections = null;
            _collectionBuild = Task.Run(() =>
            {
                var collections = Collections.LoadAll().ToArray();
                return new CollectionSnapshot { Engine = new PhrasePredictionEngine(collections),
                    BuiltInIds = collections.Where(c => c.BuiltIn).Select(c => c.Id).ToArray(),
                    AllIds = collections.Select(c => c.Id).ToArray() };
            });
        }

        private sealed class CollectionSnapshot
        {
            internal PhrasePredictionEngine Engine;
            internal string[] BuiltInIds;
            internal string[] AllIds;
        }
        private sealed class QueryBatch
        {
            internal string[] Lexical;
            internal PredictionCandidate[] Phrases;
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
            if (SuggestionVisible) _metrics.Count("dismissed");
            HideSuggestion();
        }

        private void HideSuggestion()
        {
            _displayedCompletion = null;
            _continuedWord = null;
            _popup.Hide();
            _ghost.Hide();
            _shown = new string[0];
            _cancel?.Cancel();
            // Keep current identity, suppress redisplay until the user edits/moves.
            if (!IsEnabled) _timer.Stop();
        }

        private void Clear()
        {
            _displayedCompletion = null;
            _continuedWord = null;
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
                if (_collectionBuild != null && _collectionBuild.IsCompleted)
                {
                    var ready = _collectionBuild; _collectionBuild = null;
                    if (ready.Status == TaskStatus.RanToCompletion)
                    {
                        _collections = ready.Result;
                        _collectionRevision++; Clear();
                    }
                    else { var error = ready.Exception; Logger.Warn("MatnAI тўпламлари юкланмади."); }
                }
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
                    if (_undoPending)
                    {
                        _undoPending = false;
                        if (_beforeAcceptance != null && _beforeAcceptance.SameAs(latest)) _metrics.Count("immediate_keyboard_undo");
                        _beforeAcceptance?.Dispose(); _beforeAcceptance = null;
                    }
                    if (latest == null) { Clear(); return; }
                    if (_current == null || !_current.SameAs(latest))
                    {
                        string retained;
                        if (_ghost.Visible && _current != null &&
                            _current.TryContinueSuggestion(latest, _displayedCompletion, out retained))
                        {
                            var continuations = new List<string>();
                            var origins = new Dictionary<string, string>(StringComparer.Ordinal);
                            for (int index = 0; index < _shown.Length; index++)
                            {
                                string next;
                                string candidate = index == 0 ? _displayedCompletion : _shown[index];
                                if (!_current.TryContinueSuggestion(latest, candidate, out next)) continue;
                                continuations.Add(next);
                                string origin;
                                if (_provenance.TryGetValue(_shown[index], out origin)) origins[next] = origin;
                            }
                            _cancel?.Cancel();
                            _current.Dispose();
                            _current = WordCompletionContext.Capture(_app);
                            if (_current == null || !_current.SameAs(latest)) { Clear(); return; }
                            _shown = continuations.Distinct(StringComparer.Ordinal).ToArray();
                            _provenance = origins;
                            _displayedCompletion = retained;
                            _continuedWord = retained;
                            _requested = _current; // Keep the stable visible choice; don't rerank every letter.
                        }
                        else
                        {
                            if (SuggestionVisible) _metrics.Count("typed_past_or_caret_moved");
                            // An incompatible edit must discard the old surface and
                            // request new evidence rather than insert a stale phrase.
                            string continued = null;
                            if (_current != null && _current.CanContinueAt(latest))
                            {
                                string previous = _ghost.Visible ? _shown.FirstOrDefault() : _continuedWord;
                                if (WordCompletionEngine.CanContinueWord(latest.Prefix, previous))
                                    continued = previous;
                            }
                            Clear();
                            _continuedWord = continued;
                            _current = WordCompletionContext.Capture(_app);
                            _latency.Restart();
                            if (_current == null) return;
                        }
                    }
                }
                if (SuggestionVisible && _current != null)
                    PositionPresentation();
                if (_query != null)
                {
                    if (!_query.IsCompleted) return; // At most one active query.
                    var finished = _query; _query = null;
                    bool valid = !_cancel.IsCancellationRequested && _queryRevision == _collectionRevision && _current != null && _current.SameAs(_queryContext);
                    _queryContext.Dispose(); _queryContext = null;
                    _cancel.Dispose(); _cancel = null;
                    if (valid && finished.Status == TaskStatus.RanToCompletion)
                    {
                        // Worker generates bounded proposals only. The SAME runtime morphology
                        // engine validates them here, on the dictionary's owning UI thread.
                        var lexical = WordCompletionEngine.ValidateCandidates(finished.Result.Lexical,
                            candidate => _dictionary.Contains(candidate) || _morphology.Analyze(candidate).IsValidInflectedForm,
                            _settings.MaxPredictions).Select(candidate => PredictionCasing.Apply(
                                _current.PrecedingContext, _current.Prefix, candidate));
                        _provenance.Clear();
                        foreach (var candidate in finished.Result.Phrases)
                            if (WordCompletionContext.IsSafeCompletion(candidate.FullCompletion) && candidate.Provenance.Count > 0)
                                _provenance[candidate.FullCompletion] = candidate.Provenance[0].CollectionId;
                        _shown = finished.Result.Phrases.Select(p => p.FullCompletion).Where(_provenance.ContainsKey)
                            .Concat(lexical).Distinct(StringComparer.Ordinal).Take(_settings.MaxPredictions).ToArray();
                        if (_shown.Length > 0) PresentPrimarySuggestion();
                        else { _metrics.Count("no_candidate"); HideSuggestion(); }
                    }
                    else if (finished.IsFaulted) Logger.Warn("MatnAI таклифи тайёрланмади.");
                    return;
                }
                if (_engine == null || _current == null ||
                    (_current.Prefix.Length > 0 && _current.Prefix.Count(char.IsLetter) < _settings.MinPredictionLength) ||
                    _requested == _current) return;
                if (!IsEnabled) { _timer.Stop(); return; }
                _requested = _current;
                _queryContext = WordCompletionContext.Capture(_app);
                if (_queryContext == null) return;
                string prefix = _queryContext.Prefix;
                var engine = _engine;
                var preferences = _settings.MatnAiLearningConsent ? _preferences.DictionarySnapshot() : null;
                var acceptanceCounts = _settings.MatnAiLearningConsent ? _preferences.Snapshot() : null;
                var phraseEngine = _collections?.Engine;
                var activeIds = GetActiveCollectionIds().Concat(_collections?.BuiltInIds ?? new string[0]).Distinct().ToArray();
                var request = new PredictionRequest(_queryContext.PrecedingContext, prefix, ScriptType.Unknown,
                    activeIds, ++_nextRequest, 10, acceptanceCounts);
                _queryRevision = _collectionRevision;
                _metrics.Count("queries");
                string continuedWord = _continuedWord;
                _cancel = new CancellationTokenSource();
                var token = _cancel.Token;
                _query = Task.Run(async () =>
                {
                    var phrases = phraseEngine == null ? new PredictionCandidate[0] :
                        await phraseEngine.PredictAsync(request, token).ConfigureAwait(false);
                    return new QueryBatch { Phrases = phrases, Lexical = engine.Complete(prefix, 10, token,
                        deferMorphologyValidation: true, acceptanceCounts: preferences, continuedWord: continuedWord) };
                }, token);
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
                ? null : GetTail(_shown[0]);
            if (tail == null)
            {
                _ghost.Hide();
                return;
            }
            try
            {
                var anchor = _current.GetCaretAnchor(_app);
                if (_ghost.Present(tail, anchor, new WindowOwner(_current.WindowHandle)))
                {
                    _displayedCompletion = _current.Prefix + _ghost.SuggestionTail;
                    _metrics.Count("shown"); _metrics.Latency(_latency.ElapsedMilliseconds);
                }
                else _metrics.Count("not_displayed");
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
                    string tail = _shown.Length == 0 ? null : GetTail(_shown[0]);
                    if (tail == null || !_ghost.Present(tail, anchor, owner)) _ghost.Hide();
                    else _displayedCompletion = _current.Prefix + _ghost.SuggestionTail;
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                _popup.Hide();
                _ghost.Hide();
            }
        }

        private string GetTail(string candidate)
        {
            if (_current == null || candidate == null || candidate.Length <= _current.Prefix.Length ||
                !candidate.StartsWith(_current.Prefix, StringComparison.Ordinal)) return null;
            return candidate.Substring(_current.Prefix.Length);
        }

        public void ShowAlternatives()
        {
            if (!CanOpenAlternatives) return;
            try
            {
                _ghost.Hide();
                var anchor = _current.GetCaretAnchor(_app);
                _popup.Present(_shown, anchor, new WindowOwner(_current.WindowHandle));
                if (_popup.Visible) _metrics.Count("alternatives_opened");
            }
            catch (System.Runtime.InteropServices.COMException) { _popup.Hide(); }
        }

        public void AcceptSelected()
        {
            if (!CanAcceptSuggestion) return;
            Accept(_popup.Visible ? _popup.Selected : _displayedCompletion);
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
                var before = WordCompletionContext.Capture(_app);
                bool inserted;
                try { inserted = _current.TryInsert(_app, word); }
                catch { before?.Dispose(); throw; }
                if (inserted)
                {
                    _beforeAcceptance?.Dispose(); _beforeAcceptance = before;
                    _lastAcceptTimestamp = Stopwatch.GetTimestamp();
                    _metrics.Count("accepted");
                    if (_settings.MatnAiLearningConsent)
                    {
                        string source = _provenance.Where(p => p.Key == word || p.Key.StartsWith(word + " ", StringComparison.Ordinal))
                            .Select(p => p.Value).FirstOrDefault() ?? AcceptanceStore.DictionaryScope;
                        try { _preferences.Record(source, word); } catch { Logger.Warn("MatnAI шахсий созламаларини сақлаб бўлмади."); }
                    }
                    _ghost.Hide(); _popup.Hide();
                    try { _metrics.Save(); } catch { Logger.Warn("MatnAI ҳисоботини сақлаб бўлмади."); }
                }
                else { _metrics.Count("insertion_rejected"); before?.Dispose(); }
            }
            catch (Exception ex) { Logger.Warn("MatnAI таклифини қўшиб бўлмади: " + ex.GetType().Name); }
            finally { _autoCorrect.IsEnabled = autoCorrect; HideSuggestion(); }
        }
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            HotkeyManager.KeyObserved -= OnKeyObserved;
            _beforeAcceptance?.Dispose(); _beforeAcceptance = null;
            try { _metrics.Save(); } catch { }
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
