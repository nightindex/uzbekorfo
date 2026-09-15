using System;
using System.Collections.Generic;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Script-normalizing facade over the validated, data-driven suffix parser.
    /// </summary>
    public class UzbekMorphAnalyzer
    {
        private readonly IDictionaryService _dictionary;
        private readonly ITransliterator _transliterator;
        private UzbekSuffixParser _suffixParser;

        public int LoadedSuffixCount { get; private set; }

        public UzbekMorphAnalyzer(IDictionaryService dictionary, ITransliterator transliterator)
        {
            _dictionary = dictionary ?? throw new ArgumentNullException(nameof(dictionary));
            _transliterator = transliterator ?? throw new ArgumentNullException(nameof(transliterator));
        }

        /// <summary>
        /// Loads and validates the complete morphology rule set. A malformed or
        /// incompatible file disables productive acceptance instead of partially
        /// applying unsafe rules.
        /// </summary>
        public void LoadSuffixes(string suffixesPath)
        {
            _suffixParser = null;
            LoadedSuffixCount = 0;
            try
            {
                MorphologyRuleSet rules = MorphologyRuleSetLoader.Load(suffixesPath);
                _suffixParser = new UzbekSuffixParser(_dictionary, rules);
                LoadedSuffixCount = rules.Suffixes.Count;
                Logger.Info($"Morphology rules loaded: {LoadedSuffixCount} suffixes");
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to load morphology rules; productive suffix checking is disabled", ex);
            }
        }

        public MorphAnalysis Analyze(string word)
        {
            if (string.IsNullOrWhiteSpace(word))
                return new MorphAnalysis { OriginalWord = word ?? string.Empty };

            string normalized = TextHelper.NormalizeWord(word);
            ScriptType script = _transliterator.DetectScript(word);
            string cyrillic = script == ScriptType.Latin
                ? _transliterator.ToCyrillic(normalized)
                : normalized;

            MorphAnalysis result = _suffixParser != null
                ? _suffixParser.Parse(cyrillic)
                : new MorphAnalysis { Root = cyrillic };
            result.OriginalWord = word;
            result.NormalizedWord = normalized;
            result.Script = script;
            return result;
        }

        public bool IsValidInflectedForm(string word)
        {
            return !string.IsNullOrWhiteSpace(word) && Analyze(word).IsValidInflectedForm;
        }

        /// <summary>
        /// Returns the reviewed lemma when available, otherwise the resolved
        /// dictionary root retained for compatibility with definition lookup.
        /// </summary>
        public string GetDictionaryRoot(MorphAnalysis analysis)
        {
            if (analysis == null) return null;
            if (!string.IsNullOrWhiteSpace(analysis.Lemma) && _dictionary.Contains(analysis.Lemma))
                return analysis.Lemma;
            if (!string.IsNullOrWhiteSpace(analysis.Root) && _dictionary.Contains(analysis.Root))
                return analysis.Root;
            return analysis.Root;
        }

        public string GetExpectedDativeSuffix(string root)
        {
            return _suffixParser?.GetExpectedAllomorph("dative", root) ?? "га";
        }

        public string GetExpectedLocativeSuffix(string root)
        {
            return _suffixParser?.GetExpectedAllomorph("locative", root) ?? "да";
        }

        public string GetExpectedAblativeSuffix(string root)
        {
            return _suffixParser?.GetExpectedAllomorph("ablative", root) ?? "дан";
        }

        public bool HasMissingPossessiveMutation(string root, string suffixId)
        {
            return _suffixParser != null &&
                   _suffixParser.RequiresConfiguredMutation(root, suffixId);
        }
    }
}
