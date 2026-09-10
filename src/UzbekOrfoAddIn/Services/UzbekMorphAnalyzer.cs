using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Morphological analyzer for Uzbek words.
    /// Performs greedy longest-suffix stripping to decompose inflected forms
    /// into root + ordered suffix chain. Works in both Cyrillic and Latin
    /// scripts by normalizing to Cyrillic internally.
    /// </summary>
    public class UzbekMorphAnalyzer
    {
        private readonly DictionaryService _dictionary;
        private readonly ITransliterator _transliterator;

        // Suffix data loaded from uzbek_suffixes.json
        private List<SuffixEntry> _allSuffixes = new List<SuffixEntry>();

        // Suffixes grouped by order for iterative stripping (outer → inner)
        // Order 6 (verb agreement) stripped first, then 5 (imperative), etc.
        private List<List<SuffixEntry>> _suffixesByOrder;

        // ---- Voicing categories for allomorph validation ----

        /// <summary>
        /// Voiceless consonants in Uzbek Cyrillic — after these, locative is -та
        /// (not -да), ablative is -тан (not -дан), dative after к is -ка.
        /// </summary>
        private static readonly HashSet<char> VoicelessConsonants = new HashSet<char>
        {
            'к', 'қ', 'п', 'т', 'с', 'ш', 'ч', 'ф', 'ҳ', 'х', 'ц'
        };

        private static readonly HashSet<char> CyrillicVowels = new HashSet<char>
        {
            'а', 'е', 'ё', 'и', 'о', 'у', 'ў', 'э', 'ю', 'я'
        };

        // Possessive consonant mutation pairs (forward: root mutation before possessive)
        private static readonly Dictionary<char, char> PossessiveMutationForward = new Dictionary<char, char>
        {
            { 'к', 'г' },
            { 'қ', 'ғ' }
        };

        // Reverse mutation (to reconstruct original root from mutated form)
        private static readonly Dictionary<char, char> PossessiveMutationReverse = new Dictionary<char, char>
        {
            { 'г', 'к' },
            { 'ғ', 'қ' }
        };

        public UzbekMorphAnalyzer(DictionaryService dictionary, ITransliterator transliterator)
        {
            _dictionary = dictionary ?? throw new ArgumentNullException(nameof(dictionary));
            _transliterator = transliterator ?? throw new ArgumentNullException(nameof(transliterator));
        }

        // =====================================================================
        //  LOADING
        // =====================================================================

        /// <summary>
        /// Loads suffix definitions from the given JSON file path.
        /// </summary>
        public void LoadSuffixes(string suffixesPath)
        {
            try
            {
                if (!File.Exists(suffixesPath))
                {
                    Logger.Warn($"Suffix file not found: {suffixesPath}");
                    return;
                }

                string json = File.ReadAllText(suffixesPath, System.Text.Encoding.UTF8);
                var serializer = new JavaScriptSerializer();
                var data = serializer.Deserialize<Dictionary<string, object>>(json);

                if (data == null || !data.ContainsKey("suffixes")) return;

                var suffixList = data["suffixes"] as object[];
                if (suffixList == null) return;

                _allSuffixes.Clear();
                foreach (var item in suffixList)
                {
                    var dict = item as Dictionary<string, object>;
                    if (dict == null) continue;

                    var entry = new SuffixEntry
                    {
                        Id = GetStr(dict, "id"),
                        Cyrillic = GetStr(dict, "cyrillic"),
                        Latin = GetStr(dict, "latin"),
                        Category = GetStr(dict, "category"),
                        AttachesTo = GetStr(dict, "attachesTo"),
                        Order = GetInt(dict, "order"),
                        AllomorphCondition = GetStr(dict, "allomorphCondition"),
                        MutatesRoot = GetStr(dict, "mutatesRoot"),
                        Person = GetStr(dict, "person"),
                        Description = GetStr(dict, "description")
                    };

                    if (!string.IsNullOrEmpty(entry.Id) && !string.IsNullOrEmpty(entry.Cyrillic))
                        _allSuffixes.Add(entry);
                }

                // Group by order descending (strip outer suffixes first)
                _suffixesByOrder = _allSuffixes
                    .GroupBy(s => s.Order)
                    .OrderByDescending(g => g.Key)
                    .Select(g => g.OrderByDescending(s => s.Cyrillic.Length).ToList())
                    .ToList();

                Logger.Info($"Морфологик анализатор: {_allSuffixes.Count} та қўшимча юкланди");
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to load suffixes", ex);
            }
        }

        // =====================================================================
        //  ANALYSIS
        // =====================================================================

        /// <summary>
        /// Performs morphological analysis on a single word.
        /// Returns the decomposition into root + suffix chain.
        /// </summary>
        public MorphAnalysis Analyze(string word)
        {
            if (string.IsNullOrWhiteSpace(word))
                return new MorphAnalysis { OriginalWord = word ?? "" };

            var result = new MorphAnalysis
            {
                OriginalWord = word,
                NormalizedWord = TextHelper.NormalizeWord(word),
                Script = _transliterator.DetectScript(word)
            };

            // Normalize to Cyrillic for analysis
            string cyrillic = result.NormalizedWord;
            if (result.Script == ScriptType.Latin)
            {
                cyrillic = _transliterator.ToCyrillic(result.NormalizedWord);
            }

            // If the whole word is in the dictionary, it's a known root with no suffixes
            if (_dictionary.Contains(cyrillic))
            {
                result.Root = cyrillic;
                result.IsKnownRoot = true;
                result.PartOfSpeech = "unknown";
                return result;
            }

            // Greedy longest-suffix stripping
            string remaining = cyrillic;
            var suffixIds = new List<string>();
            var suffixForms = new List<string>();
            bool foundVerb = false;
            bool foundNoun = false;

            if (_suffixesByOrder != null)
            {
                foreach (var orderGroup in _suffixesByOrder)
                {
                    var matched = FindLongestMatchingSuffix(remaining, orderGroup, foundVerb, foundNoun);
                    if (matched != null)
                    {
                        string suffixForm = matched.Cyrillic;
                        remaining = remaining.Substring(0, remaining.Length - suffixForm.Length);
                        suffixIds.Insert(0, matched.Id);
                        suffixForms.Insert(0, suffixForm);

                        if (matched.AttachesTo != null && matched.AttachesTo.Contains("verb"))
                            foundVerb = true;
                        if (matched.AttachesTo != null && matched.AttachesTo.Contains("noun"))
                            foundNoun = true;
                    }
                }
            }

            result.SuffixIds = suffixIds;
            result.Suffixes = suffixForms;

            // Try to validate the root
            result.Root = remaining;
            result.IsKnownRoot = IsKnownRootForm(remaining);

            // Determine POS from suffix chain
            if (foundVerb)
                result.PartOfSpeech = "verb";
            else if (foundNoun || suffixIds.Any(id => id.StartsWith("POSS_") || id == "PL" ||
                     id.StartsWith("DAT_") || id == "ACC" || id.StartsWith("LOC_") ||
                     id.StartsWith("ABL_") || id == "GEN"))
                result.PartOfSpeech = "noun";

            return result;
        }

        /// <summary>
        /// Checks if a root form is known (directly or via possessive reverse mutation).
        /// </summary>
        private bool IsKnownRootForm(string root)
        {
            if (string.IsNullOrEmpty(root) || root.Length < 2) return false;

            // Direct lookup
            if (_dictionary.Contains(root)) return true;

            // Try reverse possessive mutation: if root ends in г, try к; if ғ, try қ
            char lastChar = root[root.Length - 1];
            if (PossessiveMutationReverse.ContainsKey(lastChar))
            {
                string reversed = root.Substring(0, root.Length - 1) + PossessiveMutationReverse[lastChar];
                if (_dictionary.Contains(reversed)) return true;
            }

            return false;
        }

        /// <summary>
        /// Returns the original dictionary root, applying reverse mutation if needed.
        /// </summary>
        public string GetDictionaryRoot(MorphAnalysis analysis)
        {
            if (analysis == null || string.IsNullOrEmpty(analysis.Root)) return null;

            if (_dictionary.Contains(analysis.Root))
                return analysis.Root;

            char lastChar = analysis.Root[analysis.Root.Length - 1];
            if (PossessiveMutationReverse.ContainsKey(lastChar))
            {
                string reversed = analysis.Root.Substring(0, analysis.Root.Length - 1)
                    + PossessiveMutationReverse[lastChar];
                if (_dictionary.Contains(reversed))
                    return reversed;
            }

            return analysis.Root;
        }

        // =====================================================================
        //  ALLOMORPH VALIDATION (used by GrammarEngine rules)
        // =====================================================================

        /// <summary>
        /// For a given root, returns the correct dative suffix in Cyrillic.
        /// </summary>
        public string GetExpectedDativeSuffix(string root)
        {
            if (string.IsNullOrEmpty(root)) return "га";
            char last = root[root.Length - 1];
            if (last == 'к') return "ка";
            if (last == 'қ' || last == 'ғ') return "қа";
            return "га";
        }

        /// <summary>
        /// For a given root, returns the correct locative suffix in Cyrillic.
        /// </summary>
        public string GetExpectedLocativeSuffix(string root)
        {
            if (string.IsNullOrEmpty(root)) return "да";
            char last = root[root.Length - 1];
            return VoicelessConsonants.Contains(last) ? "та" : "да";
        }

        /// <summary>
        /// For a given root, returns the correct ablative suffix in Cyrillic.
        /// </summary>
        public string GetExpectedAblativeSuffix(string root)
        {
            if (string.IsNullOrEmpty(root)) return "дан";
            char last = root[root.Length - 1];
            return VoicelessConsonants.Contains(last) ? "тан" : "дан";
        }

        /// <summary>
        /// Checks if a possessive suffix requires root mutation (к→г, қ→ғ)
        /// that hasn't been applied.
        /// </summary>
        public bool HasMissingPossessiveMutation(string root, string suffixId)
        {
            if (string.IsNullOrEmpty(root) || string.IsNullOrEmpty(suffixId)) return false;

            // Only possessive suffixes trigger mutation
            if (!suffixId.StartsWith("POSS_")) return false;

            char last = root[root.Length - 1];

            // If the root still ends in к or қ (not mutated to г/ғ),
            // AND the original (unmutated) form is in the dictionary,
            // then mutation is missing.
            if (PossessiveMutationForward.ContainsKey(last))
            {
                // The root should have been mutated but wasn't
                return _dictionary.Contains(root);
            }

            return false;
        }

        /// <summary>
        /// Returns true if the character is a Cyrillic vowel.
        /// </summary>
        public static bool IsVowel(char c) => CyrillicVowels.Contains(char.ToLowerInvariant(c));

        /// <summary>
        /// Returns true if the character is a voiceless consonant.
        /// </summary>
        public static bool IsVoiceless(char c) => VoicelessConsonants.Contains(char.ToLowerInvariant(c));

        // =====================================================================
        //  SUFFIX MATCHING
        // =====================================================================

        private SuffixEntry FindLongestMatchingSuffix(string word, List<SuffixEntry> candidates,
            bool verbContext, bool nounContext)
        {
            if (string.IsNullOrEmpty(word) || word.Length < 3) return null;

            foreach (var suffix in candidates) // already sorted by length desc
            {
                if (word.Length <= suffix.Cyrillic.Length) continue;
                if (word.Length - suffix.Cyrillic.Length < 2) continue; // root must be ≥2 chars

                if (word.EndsWith(suffix.Cyrillic, StringComparison.Ordinal))
                {
                    return suffix;
                }
            }

            return null;
        }

        // =====================================================================
        //  JSON HELPERS
        // =====================================================================

        private static string GetStr(Dictionary<string, object> dict, string key)
        {
            if (dict.ContainsKey(key) && dict[key] != null)
                return dict[key].ToString();
            return null;
        }

        private static int GetInt(Dictionary<string, object> dict, string key)
        {
            if (dict.ContainsKey(key) && dict[key] != null)
            {
                if (dict[key] is int i) return i;
                int.TryParse(dict[key].ToString(), out int result);
                return result;
            }
            return 0;
        }
    }
}
