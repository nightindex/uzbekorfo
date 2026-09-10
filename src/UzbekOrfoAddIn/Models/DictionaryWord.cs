using System;

namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Represents a word in the user's personal custom dictionary.
    /// </summary>
    public class DictionaryWord
    {
        /// <summary>
        /// The dictionary word itself (normalized lowercase).
        /// </summary>
        public string Word { get; set; }

        /// <summary>
        /// Optional note or category (e.g., "legal term", "proper noun").
        /// </summary>
        public string Note { get; set; }

        /// <summary>
        /// When this word was added to the user dictionary.
        /// </summary>
        public DateTime AddedDate { get; set; } = DateTime.Now;

        public DictionaryWord() { }

        public DictionaryWord(string word, string note = null)
        {
            Word = word;
            Note = note;
            AddedDate = DateTime.Now;
        }

        public override string ToString() => Word;

        public override bool Equals(object obj) =>
            obj is DictionaryWord other &&
            string.Equals(Word, other.Word, StringComparison.OrdinalIgnoreCase);

        public override int GetHashCode() =>
            Word?.ToLowerInvariant().GetHashCode() ?? 0;
    }
}
