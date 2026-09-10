using System.Collections.Generic;

namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Result of morphological analysis for a single Uzbek word.
    /// Contains the root, detected suffixes, part-of-speech guess, and script info.
    /// </summary>
    public class MorphAnalysis
    {
        /// <summary>Original word as it appeared in text.</summary>
        public string OriginalWord { get; set; }

        /// <summary>Normalized (lowercased, trimmed) word used for analysis.</summary>
        public string NormalizedWord { get; set; }

        /// <summary>Script of the original word.</summary>
        public ScriptType Script { get; set; }

        /// <summary>Extracted root after suffix stripping.</summary>
        public string Root { get; set; }

        /// <summary>Ordered list of detected suffix IDs (inner → outer).</summary>
        public List<string> SuffixIds { get; set; } = new List<string>();

        /// <summary>Ordered list of detected suffix surface forms (Cyrillic).</summary>
        public List<string> Suffixes { get; set; } = new List<string>();

        /// <summary>Whether the extracted root is found in the dictionary.</summary>
        public bool IsKnownRoot { get; set; }

        /// <summary>Guessed part of speech: "noun", "verb", or "unknown".</summary>
        public string PartOfSpeech { get; set; } = "unknown";

        /// <summary>True if the word could be successfully decomposed.</summary>
        public bool IsAnalyzed => SuffixIds.Count > 0 || IsKnownRoot;
    }
}
