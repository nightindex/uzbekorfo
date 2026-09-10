using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// Core spelling engine interface — checks words, returns suggestions, highlights errors.
    /// </summary>
    public interface ISpellingEngine
    {
        /// <summary>
        /// Checks if a single word is spelled correctly.
        /// </summary>
        bool IsCorrect(string word);

        /// <summary>
        /// Returns a ranked list of correction suggestions for a misspelled word.
        /// </summary>
        List<Suggestion> GetSuggestions(string word, int maxResults = 5);

        /// <summary>
        /// Scans a Word range for spelling errors and returns all found errors.
        /// </summary>
        List<ErrorEntry> Check(Word.Range range);

        /// <summary>
        /// Applies red wavy underlines to all error ranges in the document.
        /// </summary>
        void HighlightErrors(List<ErrorEntry> errors);

        /// <summary>
        /// Removes all spelling error highlights from the document.
        /// </summary>
        void ClearHighlights(Word.Document document);
    }
}
