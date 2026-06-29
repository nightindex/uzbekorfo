namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Final decision from ranker guardrails for automatic replacement.
    /// </summary>
    public class AutoCorrectDecision
    {
        /// <summary>
        /// True when replacement is considered safe.
        /// </summary>
        public bool ShouldReplace { get; set; }

        /// <summary>
        /// Replacement word selected by ranker.
        /// </summary>
        public string SelectedReplacement { get; set; }

        /// <summary>
        /// Score of the selected replacement (0.0 - 1.0).
        /// </summary>
        public double Score { get; set; }

        /// <summary>
        /// Explainability reason code for logging.
        /// </summary>
        public string Reason { get; set; }

        /// <summary>
        /// Ranker version that produced this decision.
        /// </summary>
        public string RankerVersion { get; set; }
    }
}
