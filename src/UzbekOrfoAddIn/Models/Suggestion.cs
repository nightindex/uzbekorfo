namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// A single correction suggestion for a misspelled word.
    /// </summary>
    public class Suggestion
    {
        /// <summary>
        /// The suggested replacement text.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Confidence score (0.0 to 1.0) — higher means more likely correct.
        /// </summary>
        public double Confidence { get; set; }

        /// <summary>
        /// Edit distance from the original misspelled word (lower = closer match).
        /// </summary>
        public int EditDistance { get; set; }

        /// <summary>
        /// Deterministic ranker score used for ordering and auto-correct guardrails.
        /// </summary>
        public double RankScore { get; set; }

        /// <summary>
        /// Machine-readable reason for why this suggestion ranked highly.
        /// </summary>
        public string ReasonCode { get; set; }

        /// <summary>
        /// Version of the ranker that produced this suggestion.
        /// </summary>
        public string RankerVersion { get; set; }

        public Suggestion() { }

        public Suggestion(string text, double confidence, int editDistance = 0)
        {
            Text = text;
            Confidence = confidence;
            RankScore = confidence;
            EditDistance = editDistance;
        }

        public override string ToString() =>
            $"{Text} ({Confidence:P0})";
    }
}
