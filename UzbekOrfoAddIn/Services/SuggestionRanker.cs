using System;
using System.Collections.Generic;
using System.Linq;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Advanced word ranker with lexical similarity + local language model signals.
    /// Designed to stay deterministic, fast, and safe for VSTO real-time typing.
    /// </summary>
    public sealed class SuggestionRanker : ISuggestionRanker
    {
        private static readonly HashSet<char> ApostropheLikeChars = new HashSet<char>(new[]
        {
            '\'',
            '\u2018',
            '\u2019',
            '\u02BB',
            '\u02BC',
            '\u0060',
            '\u00B4',
            '\u02B9',
            '\uFF07'
        });

        private readonly DictionaryService _dictionary;
        private readonly LanguageModelService _languageModel;
        private readonly SettingsManager _settings;

        public string RankerVersion => "advanced-v1";

        public SuggestionRanker(
            DictionaryService dictionary,
            LanguageModelService languageModel,
            SettingsManager settings)
        {
            _dictionary = dictionary ?? throw new ArgumentNullException(nameof(dictionary));
            _languageModel = languageModel;
            _settings = settings;
        }

        public List<Suggestion> RankCandidates(
            string originalWord,
            string canonicalInputWord,
            ScriptType inputScript,
            IList<CandidateMatch> candidates,
            SuggestionContext context,
            int maxResults)
        {
            if (string.IsNullOrWhiteSpace(canonicalInputWord) || candidates == null || candidates.Count == 0)
                return new List<Suggestion>();

            int maxDistance = 1;
            for (int i = 0; i < candidates.Count; i++)
            {
                if (candidates[i] != null && candidates[i].Distance > maxDistance)
                    maxDistance = candidates[i].Distance;
            }

            string previous = context != null ? context.PreviousWord : null;
            string next = context != null ? context.NextWord : null;

            var ranked = candidates
                .Where(c => c != null && !string.IsNullOrWhiteSpace(c.Word))
                .GroupBy(c => c.Word, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderBy(c => c.Distance).First())
                .Where(c => !string.Equals(c.Word, canonicalInputWord, StringComparison.OrdinalIgnoreCase))
                .Select(c =>
                {
                    var feature = ComputeScore(
                        originalWord,
                        canonicalInputWord,
                        c.Word,
                        c.Distance,
                        maxDistance,
                        inputScript,
                        previous,
                        next);

                    return new Suggestion(c.Word, feature.Score, c.Distance)
                    {
                        RankScore = feature.Score,
                        Confidence = feature.Score,
                        ReasonCode = feature.ReasonCode,
                        RankerVersion = RankerVersion
                    };
                })
                .OrderByDescending(s => s.RankScore)
                .ThenBy(s => s.EditDistance)
                .ThenBy(s => s.Text, StringComparer.OrdinalIgnoreCase)
                .Take(Math.Max(1, maxResults))
                .ToList();

            return ranked;
        }

        public AutoCorrectDecision BuildAutoCorrectDecision(
            string originalWord,
            List<Suggestion> rankedSuggestions,
            SuggestionContext context,
            string autoCorrectMode,
            double minAutoCorrectScore)
        {
            var none = new AutoCorrectDecision
            {
                ShouldReplace = false,
                SelectedReplacement = null,
                Score = 0.0,
                Reason = "NO_CANDIDATE",
                RankerVersion = RankerVersion
            };

            if (rankedSuggestions == null || rankedSuggestions.Count == 0)
                return none;

            var best = rankedSuggestions[0];
            if (best == null || string.IsNullOrWhiteSpace(best.Text))
                return none;

            string normalized = TextHelper.NormalizeWord(originalWord);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                none.Reason = "EMPTY_INPUT";
                return none;
            }

            if (TextHelper.ShouldSkipWord(normalized))
            {
                none.Reason = "SKIP_TOKEN";
                return none;
            }

            if (best.EditDistance > 2)
            {
                none.Reason = "DISTANCE_TOO_HIGH";
                return none;
            }

            if (ContainsMixedScripts(normalized))
            {
                none.Reason = "MIXED_SCRIPT_GUARD";
                return none;
            }

            if (IsAmbiguousApostrophe(normalized, best.Text))
            {
                none.Reason = "APOSTROPHE_AMBIGUOUS";
                return none;
            }

            string mode = NormalizeMode(autoCorrectMode);
            double requiredScore = ResolveRequiredScore(mode, best.EditDistance, minAutoCorrectScore);
            if (best.RankScore < requiredScore)
            {
                none.Reason = "SCORE_BELOW_THRESHOLD";
                return none;
            }

            double requiredMargin = ResolveRequiredMargin(mode);
            double margin = 1.0;
            if (rankedSuggestions.Count > 1 && rankedSuggestions[1] != null)
                margin = best.RankScore - rankedSuggestions[1].RankScore;

            if (margin < requiredMargin)
            {
                none.Reason = "LOW_WINNER_MARGIN";
                return none;
            }

            return new AutoCorrectDecision
            {
                ShouldReplace = true,
                SelectedReplacement = best.Text,
                Score = best.RankScore,
                Reason = best.ReasonCode ?? "ADVANCED_OK",
                RankerVersion = RankerVersion
            };
        }

        public void RecordAcceptedCorrection(string sourceWord, string replacementWord, SuggestionContext context)
        {
            if (_languageModel == null) return;
            if (_settings != null && !_settings.LanguageModelEnabled) return;
            _languageModel.RecordAcceptedCorrection(sourceWord, replacementWord, context);
        }

        private ScoreFeature ComputeScore(
            string originalWord,
            string inputCanonical,
            string candidateCanonical,
            int distance,
            int maxDistance,
            ScriptType inputScript,
            string previousWord,
            string nextWord)
        {
            int inputLen = inputCanonical.Length;
            int candLen = candidateCanonical.Length;
            int maxLen = Math.Max(inputLen, candLen);
            int minLen = Math.Max(1, Math.Min(inputLen, candLen));

            double distanceScore = 1.0 - ((double)distance / (maxLen + 1.0));
            double prefixScore = (double)CommonPrefixLength(inputCanonical, candidateCanonical) / minLen;
            double suffixScore = (double)CommonSuffixLength(inputCanonical, candidateCanonical) / minLen;
            double diceScore = DiceCoefficient(inputCanonical, candidateCanonical);
            double edgeScore = (inputCanonical[0] == candidateCanonical[0] ? 0.6 : 0.0) +
                               (inputCanonical[inputCanonical.Length - 1] == candidateCanonical[candidateCanonical.Length - 1] ? 0.4 : 0.0);

            double lengthPenalty = (double)Math.Abs(inputLen - candLen) / Math.Max(1, maxLen);
            double distancePenalty = maxDistance > 0 ? (double)distance / maxDistance : 0.0;
            double phoneticBoost = TextHelper.IsPhoneticallySimilar(inputCanonical, candidateCanonical) ? 0.08 : 0.0;
            double transposeBoost = IsSingleAdjacentTransposition(inputCanonical, candidateCanonical) ? 0.07 : 0.0;

            double apostropheBoost = ScoreApostropheSignal(originalWord, candidateCanonical);
            double mixedScriptPenalty = ContainsMixedScripts(originalWord) ? 0.10 : 0.0;
            double translitPenalty = ResolveTranslitAmbiguityPenalty(inputScript, originalWord, candidateCanonical);

            double priorScore = 0.0;
            double contextScore = 0.0;
            double userHistoryBoost = 0.0;
            if (_languageModel != null && (_settings == null || _settings.LanguageModelEnabled))
            {
                priorScore = _languageModel.GetUnigramPriorScore(candidateCanonical);
                contextScore = _languageModel.GetContextScore(previousWord, candidateCanonical, nextWord);
                userHistoryBoost = _languageModel.GetAcceptedCorrectionBoost(originalWord, candidateCanonical);
            }

            double score =
                (0.30 * distanceScore) +
                (0.16 * diceScore) +
                (0.11 * prefixScore) +
                (0.08 * suffixScore) +
                (0.08 * edgeScore) +
                (0.16 * priorScore) +
                (0.07 * contextScore) +
                phoneticBoost +
                transposeBoost +
                apostropheBoost +
                userHistoryBoost -
                (0.06 * lengthPenalty) -
                (0.04 * distancePenalty) -
                mixedScriptPenalty -
                translitPenalty;

            score = Clamp01(score);
            return new ScoreFeature
            {
                Score = score,
                ReasonCode = ResolveReasonCode(distance, priorScore, contextScore, apostropheBoost, userHistoryBoost)
            };
        }

        private static string ResolveReasonCode(
            int distance,
            double priorScore,
            double contextScore,
            double apostropheBoost,
            double historyBoost)
        {
            if (historyBoost >= 0.05) return "USER_ACCEPTED_HISTORY";
            if (contextScore >= 0.24) return "CONTEXT_BIGRAM";
            if (apostropheBoost > 0.03) return "APOSTROPHE_FIX";
            if (priorScore >= 0.55 && distance <= 1) return "HIGH_PRIOR_LOW_DISTANCE";
            return "SIMILARITY_MATCH";
        }

        private static double ResolveRequiredScore(string mode, int editDistance, double minAutoCorrectScore)
        {
            double floor;
            if (editDistance <= 1) floor = 0.82;
            else if (editDistance == 2) floor = 0.93;
            else floor = 0.98;

            floor = Math.Max(floor, minAutoCorrectScore <= 0 ? floor : minAutoCorrectScore);

            if (mode == "balanced") floor -= 0.06;
            else if (mode == "aggressive") floor -= 0.12;

            return Clamp01(floor);
        }

        private static double ResolveRequiredMargin(string mode)
        {
            if (mode == "balanced") return 0.09;
            if (mode == "aggressive") return 0.06;
            return 0.13; // conservative
        }

        private static string NormalizeMode(string mode)
        {
            if (string.IsNullOrWhiteSpace(mode)) return "conservative";
            mode = mode.Trim().ToLowerInvariant();
            if (mode == "balanced" || mode == "aggressive")
                return mode;
            return "conservative";
        }

        private static double ScoreApostropheSignal(string inputWord, string candidateCanonical)
        {
            if (string.IsNullOrWhiteSpace(inputWord) || string.IsNullOrWhiteSpace(candidateCanonical))
                return 0.0;

            bool inputHasApostrophe = inputWord.Any(IsApostropheLike);
            bool candHasApostrophe = candidateCanonical.Any(IsApostropheLike);

            if (!inputHasApostrophe && !candHasApostrophe)
                return 0.0;

            if (inputHasApostrophe && candHasApostrophe)
                return 0.05;

            // Mild reward when candidate introduces likely Uzbek apostrophe form
            // and input had apostrophe confusion symbols.
            if (inputHasApostrophe && !candHasApostrophe) return -0.03;
            if (!inputHasApostrophe && candHasApostrophe) return 0.01;
            return 0.0;
        }

        private static double ResolveTranslitAmbiguityPenalty(ScriptType inputScript, string inputWord, string candidateCanonical)
        {
            if (inputScript == ScriptType.Mixed) return 0.10;
            if (string.IsNullOrWhiteSpace(inputWord) || string.IsNullOrWhiteSpace(candidateCanonical)) return 0.0;

            // Penalize script-mixed candidate surface forms.
            if (ContainsMixedScripts(candidateCanonical)) return 0.06;
            return 0.0;
        }

        private static bool ContainsMixedScripts(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return false;
            bool hasLatin = false;
            bool hasCyrillic = false;

            for (int i = 0; i < word.Length; i++)
            {
                char ch = word[i];
                if (TextHelper.IsLatin(ch)) hasLatin = true;
                else if (TextHelper.IsCyrillic(ch)) hasCyrillic = true;

                if (hasLatin && hasCyrillic) return true;
            }

            return false;
        }

        private static bool IsAmbiguousApostrophe(string originalWord, string candidateWord)
        {
            if (string.IsNullOrWhiteSpace(originalWord) || string.IsNullOrWhiteSpace(candidateWord))
                return false;

            bool inputHas = originalWord.Any(IsApostropheLike);
            bool candHas = candidateWord.Any(IsApostropheLike);

            if (inputHas && !candHas) return true;

            int distinctInputMarks = CountDistinctApostropheMarks(originalWord);
            if (distinctInputMarks > 1 && candHas) return true;

            return false;
        }

        private static int CountDistinctApostropheMarks(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            var set = new HashSet<char>();
            for (int i = 0; i < text.Length; i++)
            {
                if (IsApostropheLike(text[i]))
                    set.Add(text[i]);
            }
            return set.Count;
        }

        private static bool IsApostropheLike(char ch)
        {
            return ApostropheLikeChars.Contains(ch);
        }

        private static int CommonPrefixLength(string a, string b)
        {
            int n = Math.Min(a.Length, b.Length);
            int i = 0;
            while (i < n && a[i] == b[i]) i++;
            return i;
        }

        private static int CommonSuffixLength(string a, string b)
        {
            int i = a.Length - 1;
            int j = b.Length - 1;
            int matched = 0;
            while (i >= 0 && j >= 0 && a[i] == b[j])
            {
                matched++;
                i--;
                j--;
            }
            return matched;
        }

        private static double DiceCoefficient(string a, string b)
        {
            if (a.Length < 2 || b.Length < 2)
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase) ? 1.0 : 0.0;

            var aBigrams = BuildBigramMultiset(a);
            var bBigrams = BuildBigramMultiset(b);
            int intersection = 0;

            foreach (var kv in aBigrams)
            {
                int countB = 0;
                if (bBigrams.TryGetValue(kv.Key, out countB))
                    intersection += Math.Min(kv.Value, countB);
            }

            int total = aBigrams.Values.Sum() + bBigrams.Values.Sum();
            if (total == 0) return 0.0;
            return (2.0 * intersection) / total;
        }

        private static Dictionary<string, int> BuildBigramMultiset(string s)
        {
            var map = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < s.Length - 1; i++)
            {
                string bg = s.Substring(i, 2);
                int count = 0;
                if (map.TryGetValue(bg, out count))
                    map[bg] = count + 1;
                else
                    map[bg] = 1;
            }
            return map;
        }

        private static bool IsSingleAdjacentTransposition(string a, string b)
        {
            if (a.Length != b.Length || a.Length < 2) return false;

            int firstDiff = -1;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] == b[i]) continue;
                firstDiff = i;
                break;
            }

            if (firstDiff < 0 || firstDiff >= a.Length - 1) return false;
            if (a[firstDiff] != b[firstDiff + 1] || a[firstDiff + 1] != b[firstDiff])
                return false;

            for (int i = firstDiff + 2; i < a.Length; i++)
            {
                if (a[i] != b[i]) return false;
            }

            return true;
        }

        private static double Clamp01(double value)
        {
            if (value < 0.0) return 0.0;
            if (value > 1.0) return 1.0;
            return value;
        }

        private sealed class ScoreFeature
        {
            public double Score { get; set; }
            public string ReasonCode { get; set; }
        }
    }
}
