namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Definition, spelling rule, grammatical note, and usage explanation for a word.
    /// </summary>
    public class ExplanationEntry
    {
        /// <summary>
        /// The word being explained.
        /// </summary>
        public string Word { get; set; }

        /// <summary>
        /// Dictionary definition / meaning.
        /// </summary>
        public string Definition { get; set; }

        /// <summary>
        /// Relevant spelling rule (e.g., "use ҳ instead of х in Arabic-origin words").
        /// </summary>
        public string SpellingRule { get; set; }

        /// <summary>
        /// Grammar or usage note.
        /// </summary>
        public string GrammarNote { get; set; }

        /// <summary>
        /// Example sentences demonstrating correct usage.
        /// </summary>
        public string[] Examples { get; set; }

        public override string ToString() =>
            $"{Word}: {Definition ?? SpellingRule ?? "(маълумот йўқ)"}";
    }
}
