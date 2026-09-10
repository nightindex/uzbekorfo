using System.Collections.Generic;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// In-memory store for errors found during the most recent spelling check.
    /// Shared across all ribbon buttons that need access to check results.
    /// </summary>
    public interface IErrorStore
    {
        /// <summary>
        /// All errors from the last spelling check.
        /// </summary>
        List<ErrorEntry> Errors { get; }

        /// <summary>
        /// Whether a check has been performed at least once in this session.
        /// </summary>
        bool HasBeenChecked { get; }

        /// <summary>
        /// Replaces the entire error list with new results (called after a check).
        /// </summary>
        void SetErrors(List<ErrorEntry> errors);

        /// <summary>
        /// Removes a single error from the list (after fix/ignore/add-to-dict).
        /// </summary>
        void RemoveError(ErrorEntry error);

        /// <summary>
        /// Clears all stored errors.
        /// </summary>
        void Clear();

        /// <summary>
        /// Total number of errors currently stored.
        /// </summary>
        int Count { get; }

        /// <summary>
        /// Fires when the error list changes (for UI badge updates, etc.).
        /// </summary>
        event System.Action ErrorsChanged;
    }
}
