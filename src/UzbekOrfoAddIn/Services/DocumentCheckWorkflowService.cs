using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Runs a full document/selection check workflow:
    /// clear old marks, run spelling + grammar, highlight, and update store.
    /// </summary>
    public sealed class DocumentCheckWorkflowService
    {
        private readonly SpellingEngine _spellingEngine;
        private readonly GrammarEngine _grammarEngine;
        private readonly ErrorStore _errorStore;

        public DocumentCheckWorkflowService(
            SpellingEngine spellingEngine,
            GrammarEngine grammarEngine,
            ErrorStore errorStore)
        {
            _spellingEngine = spellingEngine ?? throw new ArgumentNullException(nameof(spellingEngine));
            _errorStore = errorStore ?? throw new ArgumentNullException(nameof(errorStore));
            _grammarEngine = grammarEngine;
        }

        public DocumentCheckResult Execute(
            Word.Document document,
            Word.Range range,
            Action<bool> setSpellHighlightFlag)
        {
            var result = new DocumentCheckResult();
            if (document == null || range == null) return result;

            _spellingEngine.ClearHighlights(document);
            _grammarEngine?.ClearHighlights(document);
            _errorStore.Clear();
            setSpellHighlightFlag?.Invoke(false);

            List<ErrorEntry> spellingErrors = _spellingEngine.Check(range) ?? new List<ErrorEntry>();
            foreach (var error in spellingErrors)
                error.OriginalText = error.Range?.Text;
            _errorStore.SetErrors(spellingErrors);
            _spellingEngine.HighlightErrors(spellingErrors);
            result.SpellingErrorCount = spellingErrors.Count;
            setSpellHighlightFlag?.Invoke(result.SpellingErrorCount > 0);

            if (_grammarEngine != null)
            {
                List<ErrorEntry> grammarErrors = _grammarEngine.CheckRange(range) ?? new List<ErrorEntry>();
                foreach (var error in grammarErrors)
                    error.OriginalText = error.Range?.Text;
                _errorStore.AddErrors(grammarErrors);
                _grammarEngine.HighlightErrors(grammarErrors);
                result.GrammarErrorCount = grammarErrors.Count;
            }

            return result;
        }
    }

    public sealed class DocumentCheckResult
    {
        public int SpellingErrorCount { get; set; }
        public int GrammarErrorCount { get; set; }
        public int TotalErrorCount => SpellingErrorCount + GrammarErrorCount;
    }
}
