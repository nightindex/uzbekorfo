using System.Collections.Generic;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// Manages the main and user custom dictionaries.
    /// </summary>
    public interface IDictionaryService
    {
        /// <summary>
        /// Loads both the main dictionary and the user custom dictionary from disk.
        /// </summary>
        void Load();

        /// <summary>
        /// Persists the user custom dictionary to disk.
        /// </summary>
        void Save();

        /// <summary>
        /// Checks if a word exists in either the main or user dictionary.
        /// </summary>
        bool Contains(string word);

        /// <summary>
        /// Adds a word to the user custom dictionary.
        /// Returns the result indicating whether the word was added or already existed.
        /// </summary>
        Models.AddWordResult AddWord(string word);

        /// <summary>
        /// Removes a word from the user custom dictionary.
        /// </summary>
        void RemoveWord(string word);

        /// <summary>
        /// Returns all words in the user custom dictionary, sorted alphabetically.
        /// </summary>
        List<string> GetUserWords();

        /// <summary>
        /// Returns all words in the main dictionary.
        /// </summary>
        List<string> GetMainWords();

        /// <summary>
        /// Returns all words from both main and user dictionaries combined, sorted alphabetically.
        /// </summary>
        List<string> GetAllWords();

        /// <summary>
        /// Checks if a word exists in the main (built-in) dictionary.
        /// Uses cross-script fallback if a transliterator is available.
        /// </summary>
        bool IsMainDictionaryWord(string word);

        /// <summary>
        /// Imports words from an external file into the user dictionary.
        /// Returns the count of newly added words.
        /// </summary>
        int ImportFromFile(string filePath);

        /// <summary>
        /// Exports user dictionary words for migration.
        /// Returns count of exported words.
        /// </summary>
        int ExportForMigration(string filePath);

        /// <summary>
        /// Total word count across main + user dictionaries.
        /// </summary>
        int TotalWordCount { get; }
    }
}
