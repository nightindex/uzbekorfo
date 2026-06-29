using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates "Replace All" preparation and execution over unresolved errors.
    /// </summary>
    public sealed class ReplaceAllWorkflowService
    {
        private readonly ErrorStore _errorStore;

        public ReplaceAllWorkflowService(ErrorStore errorStore)
        {
            _errorStore = errorStore;
        }

        public ReplaceAllAnalysis Analyze()
        {
            if (_errorStore == null || _errorStore.Count == 0)
                return ReplaceAllAnalysis.EmptyStore();

            var fixable = new List<ErrorEntry>();
            foreach (var err in _errorStore.Errors.Where(er => er != null && !er.IsResolved))
            {
                // Force lazy suggestion generation so ReplaceAll uses the same
                // ranking path as Suggestions and AutoCorrect flows.
                string best = err.EnsureBestSuggestion();
                if (!string.IsNullOrEmpty(best))
                    fixable.Add(err);
            }

            if (fixable.Count == 0)
                return ReplaceAllAnalysis.NoFixable();

            return ReplaceAllAnalysis.Ready(fixable);
        }

        public int ApplyAll(Word.Document document, IReadOnlyList<ErrorEntry> fixableErrors)
        {
            if (document == null || fixableErrors == null || fixableErrors.Count == 0)
                return 0;

            int replaced = 0;
            DocumentHelper.BeginUndoRecord("Барчасини алмаштириш");
            try
            {
                // Replace from end -> start so earlier offsets stay valid.
                var sorted = fixableErrors.OrderByDescending(err => err.StartIndex).ToList();
                foreach (var error in sorted)
                {
                    try
                    {
                        Word.Range freshRange = document.Range(error.StartIndex, error.EndIndex);
                        if (freshRange == null) continue;

                        DocumentHelper.ReplaceRangeText(freshRange, error.BestSuggestion);
                        error.IsResolved = true;
                        replaced++;
                    }
                    catch { }
                }
            }
            finally
            {
                DocumentHelper.EndUndoRecord();
            }

            return replaced;
        }
    }

    public sealed class ReplaceAllAnalysis
    {
        public ReplaceAllWorkflowStatus Status { get; private set; }
        public IReadOnlyList<ErrorEntry> FixableErrors { get; private set; }
        public int FixableCount => FixableErrors == null ? 0 : FixableErrors.Count;

        public static ReplaceAllAnalysis EmptyStore() =>
            new ReplaceAllAnalysis
            {
                Status = ReplaceAllWorkflowStatus.EmptyStore,
                FixableErrors = Array.Empty<ErrorEntry>()
            };

        public static ReplaceAllAnalysis NoFixable() =>
            new ReplaceAllAnalysis
            {
                Status = ReplaceAllWorkflowStatus.NoFixable,
                FixableErrors = Array.Empty<ErrorEntry>()
            };

        public static ReplaceAllAnalysis Ready(IReadOnlyList<ErrorEntry> fixableErrors) =>
            new ReplaceAllAnalysis
            {
                Status = ReplaceAllWorkflowStatus.Ready,
                FixableErrors = fixableErrors ?? Array.Empty<ErrorEntry>()
            };
    }

    public enum ReplaceAllWorkflowStatus
    {
        EmptyStore = 0,
        NoFixable = 1,
        Ready = 2
    }
}
