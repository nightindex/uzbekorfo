using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// In-memory store for errors found during the most recent spelling check.
    /// Fires ErrorsChanged event when the list is modified, enabling UI updates.
    /// </summary>
    public class ErrorStore : IErrorStore
    {
        private readonly List<ErrorEntry> _errors = new List<ErrorEntry>();
        private readonly object _lock = new object();
        private List<ErrorEntry> _cachedSnapshot;

        /// <inheritdoc/>
        public List<ErrorEntry> Errors
        {
            get
            {
                lock (_lock)
                {
                    if (_cachedSnapshot == null)
                        _cachedSnapshot = new List<ErrorEntry>(_errors);
                    return _cachedSnapshot;
                }
            }
        }

        private void InvalidateSnapshot()
        {
            _cachedSnapshot = null;
        }

        /// <inheritdoc/>
        public bool HasBeenChecked { get; private set; }

        /// <inheritdoc/>
        public int Count
        {
            get { lock (_lock) { return _errors.Count; } }
        }

        /// <inheritdoc/>
        public event Action ErrorsChanged;

        /// <inheritdoc/>
        public void SetErrors(List<ErrorEntry> errors)
        {
            lock (_lock)
            {
                ReleaseComRanges(_errors);
                _errors.Clear();
                if (errors != null)
                    _errors.AddRange(errors);
                HasBeenChecked = true;
                InvalidateSnapshot();
            }
            ErrorsChanged?.Invoke();
        }

        /// <summary>
        /// Appends additional errors (e.g., grammar) without clearing existing ones.
        /// </summary>
        public void AddErrors(List<ErrorEntry> errors)
        {
            if (errors == null || errors.Count == 0) return;
            lock (_lock)
            {
                _errors.AddRange(errors);
                InvalidateSnapshot();
            }
            ErrorsChanged?.Invoke();
        }

        /// <summary>
        /// Returns only spelling-related errors (Spelling, Typo, Style without RuleId).
        /// </summary>
        public List<ErrorEntry> GetSpellingErrors()
        {
            lock (_lock)
            {
                return _errors.Where(e => !e.IsGrammarError).ToList();
            }
        }

        /// <summary>
        /// Returns only grammar-related errors (Grammar, Punctuation with RuleId).
        /// </summary>
        public List<ErrorEntry> GetGrammarErrors()
        {
            lock (_lock)
            {
                return _errors.Where(e => e.IsGrammarError).ToList();
            }
        }

        /// <inheritdoc/>
        public void RemoveError(ErrorEntry error)
        {
            lock (_lock)
            {
                _errors.Remove(error);
                InvalidateSnapshot();
            }
            ReleaseComRange(error);
            ErrorsChanged?.Invoke();
        }

        /// <inheritdoc/>
        public void Clear()
        {
            lock (_lock)
            {
                ReleaseComRanges(_errors);
                _errors.Clear();
                HasBeenChecked = false;
                InvalidateSnapshot();
            }
            ErrorsChanged?.Invoke();
        }

        /// <summary>
        /// Releases COM Range references held by error entries to prevent RCW leaks.
        /// </summary>
        private static void ReleaseComRanges(List<ErrorEntry> errors)
        {
            foreach (var err in errors)
                ReleaseComRange(err);
        }

        private static void ReleaseComRange(ErrorEntry err)
        {
            if (err?.Range == null) return;
            DocumentHighlightService.ClearRange(err.Range);
            try { Marshal.ReleaseComObject(err.Range); } catch { }
            err.Range = null;
        }
    }
}
