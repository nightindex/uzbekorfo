using System;
using System.Collections.Generic;
using System.Linq;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Executes a validated, data-driven Uzbek suffix grammar. No suffix IDs,
    /// category order, compatibility groups, or mutation pairs are encoded here.
    /// </summary>
    public sealed class UzbekSuffixParser
    {
        private readonly IDictionaryService _dictionary;
        private readonly ILexemeMetadataProvider _metadata;
        private readonly MorphologyRuleSet _rules;
        private readonly List<SuffixEntry> _suffixes;
        private readonly Dictionary<char, List<SuffixEntry>> _standaloneSuffixes;
        private readonly Dictionary<string, SuffixEntry> _suffixById;
        private readonly Dictionary<string, MorphologyFamilyRule> _familyById;
        private readonly Dictionary<string, RootMutationRule> _mutationById;

        public UzbekSuffixParser(IDictionaryService dictionary, MorphologyRuleSet rules)
        {
            _dictionary = dictionary ?? throw new ArgumentNullException(nameof(dictionary));
            _rules = rules ?? throw new ArgumentNullException(nameof(rules));

            List<string> errors = rules.Validate();
            if (errors.Count > 0)
                throw new ArgumentException("Invalid morphology rule set: " + string.Join("; ", errors), nameof(rules));

            _metadata = dictionary as ILexemeMetadataProvider;
            _suffixes = rules.Suffixes
                .Where(s => !s.StandaloneOnly)
                .OrderByDescending(s => s.Cyrillic.Length)
                .ThenBy(s => s.Id, StringComparer.Ordinal)
                .ToList();
            _standaloneSuffixes = rules.Suffixes.Where(s => s.StandaloneOnly)
                .GroupBy(s => s.Cyrillic[s.Cyrillic.Length - 1])
                .ToDictionary(g => g.Key, g => g.OrderByDescending(s => s.Cyrillic.Length).ToList());
            _suffixById = rules.Suffixes.ToDictionary(s => s.Id, StringComparer.Ordinal);
            _familyById = rules.Families.ToDictionary(f => f.Id, StringComparer.Ordinal);
            _mutationById = rules.RootMutations.ToDictionary(r => r.Id, StringComparer.Ordinal);
        }

        /// <summary>
        /// Parses a normalized Cyrillic word. Exact dictionary entries remain
        /// the fastest and highest-authority spelling path.
        /// </summary>
        public MorphAnalysis Parse(string cyrillicWord)
        {
            var result = new MorphAnalysis
            {
                OriginalWord = cyrillicWord ?? string.Empty,
                NormalizedWord = cyrillicWord ?? string.Empty,
                Script = ScriptType.Cyrillic,
                Root = cyrillicWord ?? string.Empty
            };

            if (string.IsNullOrWhiteSpace(cyrillicWord))
                return result;

            if (_dictionary.Contains(cyrillicWord))
            {
                result.IsKnownRoot = true;
                result.IsValidSuffixChain = true;
                ApplyLexemeMetadata(result, cyrillicWord);
                return result;
            }

            var candidates = new List<ParseCandidate>();
            // Imported complete forms cannot enter the recursive chain parser.
            if (_standaloneSuffixes.TryGetValue(cyrillicWord[cyrillicWord.Length - 1], out var standalone))
            {
                foreach (var suffix in standalone)
                {
                    if (!cyrillicWord.EndsWith(suffix.Cyrillic, StringComparison.Ordinal) ||
                        cyrillicWord.Length < suffix.Cyrillic.Length + _rules.MinimumRootLength) continue;
                    var surfaceRoot = cyrillicWord.Substring(0, cyrillicWord.Length - suffix.Cyrillic.Length);
                    var family = _familyById[suffix.Family];
                    if (TryResolveExact(surfaceRoot, family, out var root))
                        candidates.Add(BuildCandidate(cyrillicWord, root, family, new List<SuffixEntry> { suffix }));
                }
            }
            CollectCandidates(
                cyrillicWord,
                cyrillicWord,
                new List<SuffixEntry>(),
                candidates,
                new HashSet<string>(StringComparer.Ordinal),
                0);

            ParseCandidate best = candidates
                .OrderByDescending(c => c.IsValidInflection)
                .ThenByDescending(c => c.IsStructurallyValid)
                .ThenByDescending(c => c.IsPartOfSpeechCompatible)
                .ThenByDescending(c => c.Root.Length)
                .ThenBy(c => c.Suffixes.Any(s => s.StandaloneOnly))
                .ThenBy(c => c.Suffixes.Count)
                .FirstOrDefault();

            if (best == null)
                return result;

            result.Root = best.Root;
            result.Lemma = best.Lemma;
            result.IsKnownRoot = true;
            result.SuffixIds = best.Suffixes.Select(s => s.Id).ToList();
            result.Suffixes = best.Suffixes.Select(s => s.Cyrillic).ToList();
            result.PartOfSpeech = best.PartOfSpeech ?? best.Family.Id;
            result.IsValidSuffixChain = best.IsValidInflection;
            return result;
        }

        public string GetExpectedAllomorph(string groupId, string stem)
        {
            AllomorphGroupRule group = _rules.AllomorphGroups
                .FirstOrDefault(g => string.Equals(g.Id, groupId, StringComparison.Ordinal));
            if (group == null) return null;

            foreach (string suffixId in group.SuffixIds)
            {
                if (suffixId == group.FallbackSuffixId) continue;
                if (_suffixById.TryGetValue(suffixId, out SuffixEntry suffix) &&
                    ConditionMatches(suffix.AllomorphCondition, stem))
                    return suffix.Cyrillic;
            }

            return _suffixById.TryGetValue(group.FallbackSuffixId, out SuffixEntry fallback)
                ? fallback.Cyrillic
                : null;
        }

        public bool RequiresConfiguredMutation(string stem, string suffixId)
        {
            if (string.IsNullOrEmpty(stem) || string.IsNullOrEmpty(suffixId) ||
                !_suffixById.TryGetValue(suffixId, out SuffixEntry suffix) ||
                string.IsNullOrWhiteSpace(suffix.RootMutation) ||
                !_mutationById.TryGetValue(suffix.RootMutation, out RootMutationRule mutation))
                return false;

            string last = stem[stem.Length - 1].ToString();
            return mutation.FinalCharacterMap.ContainsKey(last) && _dictionary.Contains(stem);
        }

        private void CollectCandidates(
            string fullWord,
            string remaining,
            List<SuffixEntry> outerToInner,
            List<ParseCandidate> candidates,
            HashSet<string> seen,
            int depth)
        {
            if (depth >= _rules.MaxSuffixDepth || remaining.Length < _rules.MinimumRootLength + 1)
                return;

            foreach (SuffixEntry suffix in _suffixes)
            {
                if (remaining.Length < suffix.Cyrillic.Length + _rules.MinimumRootLength ||
                    !remaining.EndsWith(suffix.Cyrillic, StringComparison.Ordinal))
                    continue;

                string rootCandidate = remaining.Substring(0, remaining.Length - suffix.Cyrillic.Length);
                var nextOuterToInner = new List<SuffixEntry>(outerToInner) { suffix };
                var innerToOuter = nextOuterToInner.AsEnumerable().Reverse().ToList();
                string familyId = DetermineFamily(innerToOuter);

                if (familyId != null && _familyById.TryGetValue(familyId, out MorphologyFamilyRule family) &&
                    TryResolveDictionaryRoot(rootCandidate, family, innerToOuter[0], out RootResolution root))
                {
                    string key = root.Root + "|" + string.Join(",", innerToOuter.Select(s => s.Id));
                    if (seen.Add(key))
                        candidates.Add(BuildCandidate(fullWord, root, family, innerToOuter));
                }

                CollectCandidates(fullWord, rootCandidate, nextOuterToInner, candidates, seen, depth + 1);
            }
        }

        private ParseCandidate BuildCandidate(
            string fullWord,
            RootResolution root,
            MorphologyFamilyRule family,
            List<SuffixEntry> suffixes)
        {
            bool structureValid = IsStructureValid(suffixes, family);
            bool productive = suffixes.All(s => s.Productive &&
                (string.IsNullOrEmpty(s.RequiredRootFlags) ||
                 s.RequiredRootFlags.Any(flag => (root.Flags ?? "").IndexOf(flag) >= 0)));
            bool formsValid = IsSurfaceFormValid(fullWord, root.Root, suffixes);
            bool posCompatible = IsPartOfSpeechCompatible(root.PartOfSpeech, family);

            return new ParseCandidate
            {
                Root = root.Root,
                Lemma = root.Lemma,
                PartOfSpeech = root.PartOfSpeech,
                Suffixes = suffixes,
                Family = family,
                IsPartOfSpeechCompatible = posCompatible,
                IsStructurallyValid = structureValid,
                IsValidInflection = structureValid && productive && formsValid && posCompatible
            };
        }

        private static string DetermineFamily(IEnumerable<SuffixEntry> suffixes)
        {
            string[] families = suffixes.Select(s => s.Family).Distinct(StringComparer.Ordinal).ToArray();
            return families.Length == 1 ? families[0] : null;
        }

        private static bool IsStructureValid(List<SuffixEntry> suffixes, MorphologyFamilyRule family)
        {
            if (suffixes == null || suffixes.Count == 0 || family == null)
                return false;

            int previousStage = 0;
            foreach (SuffixEntry suffix in suffixes)
            {
                if (!family.CategoryStages.TryGetValue(suffix.Category, out int stage) ||
                    stage < previousStage)
                    return false;
                previousStage = stage;
            }

            foreach (CategoryOccurrenceLimit limit in family.OccurrenceLimits)
            {
                int count = suffixes.Count(s => limit.Categories.Contains(s.Category));
                if (count > limit.Max) return false;
            }

            foreach (CategoryRequirementRule requirement in family.Requirements)
            {
                bool trigger = suffixes.Any(s => requirement.IfAny.Contains(s.Category));
                if (trigger && !suffixes.Any(s => requirement.RequireAny.Contains(s.Category)))
                    return false;
            }

            return true;
        }

        private bool IsSurfaceFormValid(string fullWord, string root, List<SuffixEntry> suffixes)
        {
            string stem = root;
            foreach (SuffixEntry suffix in suffixes)
            {
                if (!ConditionMatches(suffix.AllomorphCondition, stem))
                    return false;
                if (IsFallbackAllomorph(suffix) && HasMatchingSpecificAllomorph(suffix, stem))
                    return false;

                stem = ApplyMutation(stem, suffix.RootMutation);
                stem += suffix.Cyrillic;
            }
            return string.Equals(stem, fullWord, StringComparison.Ordinal);
        }

        private bool HasMatchingSpecificAllomorph(SuffixEntry suffix, string stem)
        {
            AllomorphGroupRule group = _rules.AllomorphGroups.FirstOrDefault(g =>
                g.SuffixIds.Contains(suffix.Id) && g.FallbackSuffixId == suffix.Id);
            if (group == null) return false;

            return group.SuffixIds
                .Where(id => id != group.FallbackSuffixId)
                .Select(id => _suffixById[id])
                .Any(other => ConditionMatches(other.AllomorphCondition, stem));
        }

        private bool IsFallbackAllomorph(SuffixEntry suffix)
        {
            return _rules.AllomorphGroups.Any(g =>
                g.FallbackSuffixId == suffix.Id && g.SuffixIds.Contains(suffix.Id));
        }

        private string ApplyMutation(string stem, string mutationId)
        {
            if (string.IsNullOrEmpty(stem) || string.IsNullOrWhiteSpace(mutationId) ||
                !_mutationById.TryGetValue(mutationId, out RootMutationRule mutation))
                return stem;

            string last = stem[stem.Length - 1].ToString();
            if (!mutation.FinalCharacterMap.TryGetValue(last, out string replacement))
                return stem;
            return stem.Substring(0, stem.Length - 1) + replacement;
        }

        private bool TryResolveDictionaryRoot(
            string surfaceRoot,
            MorphologyFamilyRule family,
            SuffixEntry innermostSuffix,
            out RootResolution resolution)
        {
            if (TryResolveExact(surfaceRoot, family, out resolution))
                return true;

            if (string.IsNullOrWhiteSpace(innermostSuffix.RootMutation) ||
                !_mutationById.TryGetValue(innermostSuffix.RootMutation, out RootMutationRule mutation) ||
                string.IsNullOrEmpty(surfaceRoot))
                return false;

            string last = surfaceRoot[surfaceRoot.Length - 1].ToString();
            foreach (var pair in mutation.FinalCharacterMap)
            {
                if (!string.Equals(pair.Value, last, StringComparison.Ordinal)) continue;
                string unmutated = surfaceRoot.Substring(0, surfaceRoot.Length - 1) + pair.Key;
                if (TryResolveExact(unmutated, family, out resolution))
                    return true;
            }

            return false;
        }

        private bool TryResolveExact(
            string root,
            MorphologyFamilyRule family,
            out RootResolution resolution)
        {
            resolution = null;
            if (string.IsNullOrEmpty(root) || root.Length < _rules.MinimumRootLength)
                return false;

            string dictionaryWord = root + (family.RootLookupSuffix ?? string.Empty);
            if (!_dictionary.Contains(dictionaryWord)) return false;

            string lemma = dictionaryWord;
            string partOfSpeech = "unknown";
            string flags = null;
            if (_metadata != null && _metadata.TryGetLexeme(dictionaryWord, out LexemeMetadata lexeme))
            {
                flags = lexeme.HunspellFlags;
                lemma = string.IsNullOrWhiteSpace(lexeme.Lemma) ? dictionaryWord : lexeme.Lemma;
                partOfSpeech = string.IsNullOrWhiteSpace(lexeme.PartOfSpeech)
                    ? "unknown"
                    : lexeme.PartOfSpeech;
            }

            resolution = new RootResolution
            {
                Root = root,
                DictionaryWord = dictionaryWord,
                Lemma = lemma,
                PartOfSpeech = partOfSpeech,
                Flags = flags
            };
            return true;
        }

        private static bool IsPartOfSpeechCompatible(string partOfSpeech, MorphologyFamilyRule family)
        {
            if (string.IsNullOrWhiteSpace(partOfSpeech) || partOfSpeech == "unknown")
                return family.AllowUnknownRootPartOfSpeech;
            return family.AllowedRootPartOfSpeech.Contains(partOfSpeech);
        }

        private void ApplyLexemeMetadata(MorphAnalysis analysis, string word)
        {
            if (_metadata != null && _metadata.TryGetLexeme(word, out LexemeMetadata lexeme))
            {
                analysis.Lemma = lexeme.Lemma;
                analysis.PartOfSpeech = lexeme.PartOfSpeech;
            }
            else
            {
                analysis.Lemma = word;
                analysis.PartOfSpeech = "unknown";
            }
        }

        private bool ConditionMatches(string condition, string stem)
        {
            if (string.IsNullOrWhiteSpace(condition) ||
                condition.Equals("default", StringComparison.OrdinalIgnoreCase))
                return true;

            if (condition.Equals("stemEndsInVowel", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(stem)) return false;
                string last = char.ToLowerInvariant(stem[stem.Length - 1]).ToString();
                return _rules.Vowels.Contains(last);
            }

            const string prefix = "stemEndsIn:";
            if (!condition.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(stem))
                return false;

            string final = char.ToLowerInvariant(stem[stem.Length - 1]).ToString();
            return condition.Substring(prefix.Length).Split(',')
                .Select(value => value.Trim())
                .Any(value => value == final);
        }

        private sealed class RootResolution
        {
            public string Flags;
            public string Root { get; set; }
            public string DictionaryWord { get; set; }
            public string Lemma { get; set; }
            public string PartOfSpeech { get; set; }
        }

        private sealed class ParseCandidate
        {
            public string Root { get; set; }
            public string Lemma { get; set; }
            public string PartOfSpeech { get; set; }
            public List<SuffixEntry> Suffixes { get; set; }
            public MorphologyFamilyRule Family { get; set; }
            public bool IsPartOfSpeechCompatible { get; set; }
            public bool IsStructurallyValid { get; set; }
            public bool IsValidInflection { get; set; }
        }
    }
}
