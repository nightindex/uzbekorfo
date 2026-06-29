namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// A transliteration exception — a word or phrase that should not follow
    /// standard transliteration rules (e.g., proper names, brands, abbreviations).
    /// </summary>
    public class TranslitException
    {
        /// <summary>The original text in Latin script.</summary>
        public string Original { get; set; }

        /// <summary>The fixed Cyrillic replacement text.</summary>
        public string Replacement { get; set; }

        /// <summary>Category: Brand, Technical, Geographic, Name, Other.</summary>
        public string Category { get; set; }

        /// <summary>Human-readable description of the exception.</summary>
        public string Description { get; set; }

        /// <summary>Whether this exception is active. Disabled exceptions are skipped.</summary>
        public bool Enabled { get; set; } = true;

        /// <summary>Optional note (kept for backward compatibility with existing JSON).</summary>
        public string Note { get; set; }

        public TranslitException() { }

        public TranslitException(string original, string replacement,
                                  string category = "Other",
                                  string description = null,
                                  bool enabled = true,
                                  string note = null)
        {
            Original = original;
            Replacement = replacement;
            Category = category ?? "Other";
            Description = description;
            Enabled = enabled;
            Note = note;
        }

        public override string ToString() =>
            $"{Original} \u2192 {Replacement} [{Category}]";
    }
}
