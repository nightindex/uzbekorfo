namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Reviewed lexical metadata attached to a bundled dictionary entry.
    /// </summary>
    public sealed class LexemeMetadata
    {
        public string Word { get; set; }
        public string Lemma { get; set; }
        public string PartOfSpeech { get; set; }

        /// <summary>
        /// Optional Hunspell affix flags imported from the upstream dictionary.
        /// These are retained for flag-aware morphology and are not part of the
        /// plain runtime spelling-list word key.
        /// </summary>
        public string HunspellFlags { get; set; }

        /// <summary>Combine aliases without losing reviewed lexical metadata.</summary>
        public void MergeFrom(LexemeMetadata other)
        {
            if (other == null) return;
            if ((string.IsNullOrWhiteSpace(PartOfSpeech) || PartOfSpeech == "unknown") &&
                !string.IsNullOrWhiteSpace(other.PartOfSpeech) && other.PartOfSpeech != "unknown")
            {
                PartOfSpeech = other.PartOfSpeech;
                Lemma = other.Lemma;
            }
            var flags = new System.Collections.Generic.SortedSet<char>(
                ((HunspellFlags ?? "") + (other.HunspellFlags ?? "")).ToCharArray());
            HunspellFlags = flags.Count == 0 ? null : string.Concat(flags);
        }
    }
}
