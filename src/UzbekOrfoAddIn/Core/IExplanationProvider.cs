using System.Collections.Generic;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// Provides definitions, spelling rules, and grammar notes for Uzbek words.
    /// </summary>
    public interface IExplanationProvider
    {
        /// <summary>
        /// Returns the explanation entry for a word (definition, rule, examples).
        /// Returns null if no information is available.
        /// </summary>
        ExplanationEntry GetExplanation(string word);

        /// <summary>
        /// Checks if an explanation exists for the given word.
        /// </summary>
        bool HasExplanation(string word);

        /// <summary>
        /// Returns a snapshot of all explanation entries.
        /// </summary>
        List<ExplanationEntry> GetAllEntries();

        /// <summary>
        /// Adds or updates an explanation entry for a word.
        /// </summary>
        void AddOrUpdate(ExplanationEntry entry);

        /// <summary>
        /// Adds or updates multiple explanation entries and saves once.
        /// Returns number of applied entries.
        /// </summary>
        int AddOrUpdateBatch(IEnumerable<ExplanationEntry> entries);

        /// <summary>
        /// Removes the explanation entry for a word.
        /// </summary>
        void Remove(string word);

        /// <summary>
        /// Loads the explanation database from embedded resource or file.
        /// </summary>
        void Load();

        /// <summary>
        /// Persists all explanation entries to disk.
        /// </summary>
        void Save();

        /// <summary>
        /// Count of locally loaded, editable explanation entries. Built-in
        /// metadata is loaded on demand and is intentionally excluded.
        /// </summary>
        int EntryCount { get; }
    }
}
