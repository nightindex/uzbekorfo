using System.Collections.Generic;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// Transliteration service for converting between Uzbek Latin and Cyrillic scripts.
    /// </summary>
    public interface ITransliterator
    {
        /// <summary>
        /// Converts Latin Uzbek text to Cyrillic script.
        /// </summary>
        string ToCyrillic(string latinText);

        /// <summary>
        /// Converts Cyrillic Uzbek text to Latin script.
        /// </summary>
        string ToLatin(string cyrillicText);

        /// <summary>
        /// Detects the dominant script in the given text.
        /// </summary>
        ScriptType DetectScript(string text);

        /// <summary>
        /// Returns the list of transliteration exceptions.
        /// </summary>
        List<TranslitException> GetExceptions();

        /// <summary>
        /// Adds or updates a transliteration exception.
        /// </summary>
        void AddException(TranslitException exception);

        /// <summary>
        /// Removes a transliteration exception.
        /// </summary>
        void RemoveException(string originalText);

        /// <summary>
        /// Loads exceptions from disk.
        /// </summary>
        void LoadExceptions();

        /// <summary>
        /// Saves exceptions to disk.
        /// </summary>
        void SaveExceptions();
    }
}
