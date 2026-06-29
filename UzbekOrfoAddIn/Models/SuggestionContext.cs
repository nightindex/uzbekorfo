namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Optional lexical context around a misspelled token.
    /// Used by advanced suggestion ranking and auto-correct guardrails.
    /// </summary>
    public class SuggestionContext
    {
        /// <summary>
        /// Previous neighboring word (if available).
        /// </summary>
        public string PreviousWord { get; set; }

        /// <summary>
        /// Next neighboring word (if available).
        /// </summary>
        public string NextWord { get; set; }

        /// <summary>
        /// Optional raw snippet around the token (for diagnostics).
        /// </summary>
        public string SurroundingText { get; set; }

        /// <summary>
        /// Optional script hint extracted from surrounding text.
        /// </summary>
        public ScriptType ScriptHint { get; set; } = ScriptType.Unknown;

        /// <summary>
        /// Indicates this context came from real-time auto-correct path.
        /// </summary>
        public bool IsAutoCorrectFlow { get; set; }
    }
}
