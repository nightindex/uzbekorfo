namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Represents a single grammar rule loaded from grammar_rules.json.
    /// </summary>
    public class GrammarRule
    {
        /// <summary>Unique identifier for the rule (e.g. "DAT_ALLOMORPH").</summary>
        public string RuleId { get; set; }

        /// <summary>Category: morphological, agreement, punctuation, style.</summary>
        public string Category { get; set; }

        /// <summary>Priority (lower = higher priority, checked first).</summary>
        public int Priority { get; set; }

        /// <summary>Whether the rule is currently active.</summary>
        public bool IsEnabled { get; set; }

        /// <summary>Severity string mapping to ErrorSeverity enum.</summary>
        public string Severity { get; set; }

        /// <summary>Human-readable description in Uzbek.</summary>
        public string DescriptionUz { get; set; }

        /// <summary>Human-readable description in English (for logging/debug).</summary>
        public string DescriptionEn { get; set; }
    }
}
