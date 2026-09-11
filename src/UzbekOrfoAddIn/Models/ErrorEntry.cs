using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Represents a single spelling error found during a check.
    /// </summary>
    public class ErrorEntry
    {
        /// <summary>
        /// The misspelled word.
        /// </summary>
        public string Word { get; set; }

        /// <summary>Exact checked text, including punctuation and whitespace, for stale-result checks.</summary>
        public string OriginalText { get; set; }

        /// <summary>
        /// Start character index within the paragraph.
        /// </summary>
        public int StartIndex { get; set; }

        /// <summary>
        /// End character index within the paragraph.
        /// </summary>
        public int EndIndex { get; set; }

        /// <summary>
        /// 1-based paragraph number in the document.
        /// </summary>
        public int ParagraphIndex { get; set; }

        /// <summary>
        /// Surrounding text snippet for context display.
        /// </summary>
        public string Context { get; set; }

        // --- Lazy suggestion loading ---
        private bool _suggestionsLoaded;
        private List<Suggestion> _suggestions;

        /// <summary>
        /// Optional provider for lazy suggestion loading.
        /// Set by SpellingEngine to defer expensive computation until user needs them.
        /// </summary>
        public Func<string, List<Suggestion>> SuggestionsProvider { get; set; }

        /// <summary>
        /// Ranked list of correction suggestions.
        /// Lazily computed via SuggestionsProvider on first access.
        /// </summary>
        public List<Suggestion> Suggestions
        {
            get
            {
                if (!_suggestionsLoaded && SuggestionsProvider != null)
                {
                    try { _suggestions = SuggestionsProvider(Word); }
                    catch { _suggestions = new List<Suggestion>(); }
                    _suggestionsLoaded = true;
                }
                return _suggestions ?? (_suggestions = new List<Suggestion>());
            }
            set
            {
                _suggestions = value;
                _suggestionsLoaded = true;
            }
        }

        /// <summary>
        /// Word Interop Range pointing to this error's location in the document.
        /// </summary>
        public Word.Range Range { get; set; }

        /// <summary>
        /// Severity level of the error.
        /// </summary>
        public ErrorSeverity Severity { get; set; } = ErrorSeverity.Spelling;

        /// <summary>
        /// Grammar rule ID (e.g., "DAT_ALLOMORPH"). Null for spelling errors.
        /// </summary>
        public string RuleId { get; set; }

        /// <summary>
        /// Human-readable error description/message (e.g., grammar rule explanation).
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// Grammar rule category (morphological, agreement, punctuation, style). Null for spelling.
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// Whether this is a grammar error (as opposed to a spelling error).
        /// </summary>
        public bool IsGrammarError => RuleId != null;

        /// <summary>
        /// Whether this error has been resolved (replaced, ignored, or added to dict).
        /// </summary>
        public bool IsResolved { get; set; }

        /// <summary>
        /// Timestamp when the error was detected.
        /// </summary>
        public DateTime DetectedAt { get; set; } = DateTime.Now;

        /// <summary>
        /// Returns the best (top-ranked) suggestion, or null if not yet loaded.
        /// Does NOT trigger lazy loading — safe for bulk grid display.
        /// </summary>
        public string BestSuggestion
        {
            get
            {
                if (_suggestions != null && _suggestions.Count > 0)
                    return _suggestions[0].Text;
                return null;
            }
        }

        /// <summary>
        /// Ensures suggestions are loaded (triggers lazy load if needed) and returns best.
        /// </summary>
        public string EnsureBestSuggestion()
        {
            var s = Suggestions; // triggers lazy load
            return BestSuggestion;
        }

        public override string ToString() =>
            $"{Word} \u2192 {BestSuggestion ?? "?"} (\u00A7{ParagraphIndex})";
    }
}
