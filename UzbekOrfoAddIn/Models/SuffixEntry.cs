namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Represents a single suffix entry loaded from uzbek_suffixes.json.
    /// </summary>
    public class SuffixEntry
    {
        /// <summary>Unique suffix ID (e.g. "DAT_GA", "PL").</summary>
        public string Id { get; set; }

        /// <summary>Cyrillic surface form.</summary>
        public string Cyrillic { get; set; }

        /// <summary>Latin surface form.</summary>
        public string Latin { get; set; }

        /// <summary>Functional category (e.g. "case_dative", "plural", "possessive").</summary>
        public string Category { get; set; }

        /// <summary>What part of speech this suffix attaches to.</summary>
        public string AttachesTo { get; set; }

        /// <summary>Agglutination order (1 = closest to root, higher = outer).</summary>
        public int Order { get; set; }

        /// <summary>Allomorph condition string, e.g. "stemEndsIn:к" or "default".</summary>
        public string AllomorphCondition { get; set; }

        /// <summary>Root mutation rule for possessive suffixes, e.g. "к→г,қ→ғ".</summary>
        public string MutatesRoot { get; set; }

        /// <summary>Person/number for agreement and possessive suffixes.</summary>
        public string Person { get; set; }

        /// <summary>Human-readable description in Uzbek.</summary>
        public string Description { get; set; }
    }
}
