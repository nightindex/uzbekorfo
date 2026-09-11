using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Core spelling engine — scans Word documents for Uzbek spelling errors,
    /// generates suggestions, and highlights mistakes with red wavy underlines.
    /// </summary>
    public class SpellingEngine : ISpellingEngine
    {
        private readonly DictionaryService _dictionary;

        public SpellingEngine(DictionaryService dictionary)
        {
            _dictionary = dictionary ?? throw new ArgumentNullException(nameof(dictionary));
        }

        // =====================================================================
        //  SINGLE-WORD CHECK
        // =====================================================================

        /// <inheritdoc/>
        public bool IsCorrect(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return true;

            var normalized = TextHelper.NormalizeWord(word);
            if (TextHelper.ShouldSkipWord(normalized)) return true;

            return _dictionary.Contains(normalized);
        }

        // =====================================================================
        //  SUGGESTIONS
        // =====================================================================

        /// <inheritdoc/>
        /// <remarks>
        /// Script-aware: suggestions from the (Cyrillic) dictionary are automatically
        /// converted to the user's input script (Latin or Cyrillic).
        /// </remarks>
        public List<Suggestion> GetSuggestions(string word, int maxResults = 5)
        {
            var normalized = TextHelper.NormalizeWord(word);
            if (string.IsNullOrEmpty(normalized))
                return new List<Suggestion>();

            // Detect the script the user typed in
            var inputScript = _dictionary.DetectWordScript(normalized);

            // If Latin input, transliterate to Cyrillic for dictionary search
            string searchWord = normalized;
            if (inputScript == Models.ScriptType.Latin)
            {
                string cyrForm = _dictionary.ConvertSuggestionToScript(normalized, Models.ScriptType.Cyrillic);
                if (!string.IsNullOrEmpty(cyrForm)) searchWord = cyrForm;
            }

            int maxDistance = GetAdaptiveMaxDistance(searchWord);
            int minPool = maxResults <= 1 ? 24 : 60;
            int candidatePool = Math.Max(minPool, maxResults * 12);

            var candidates = _dictionary.FindSimilarWords(
                searchWord,
                maxDistance: maxDistance,
                maxResults: candidatePool);

            foreach (var extra in GenerateDirectVariantCandidates(searchWord, maxDistance))
                candidates.Add(extra);

            var ranked = candidates
                .GroupBy(c => c.Word, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderBy(c => c.Distance).First())
                .Where(c => !string.Equals(c.Word, searchWord, StringComparison.OrdinalIgnoreCase))
                .Select(c => new
                {
                    c.Word,
                    c.Distance,
                    Score = ScoreCandidate(searchWord, c.Word, c.Distance, maxDistance)
                })
                .OrderByDescending(x => x.Score)
                .ThenBy(x => x.Distance)
                .ThenBy(x => x.Word)
                .ToList();

            var output = new List<Suggestion>(maxResults);
            var seenDisplay = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in ranked)
            {
                string displayWord = item.Word;
                if (inputScript == Models.ScriptType.Latin)
                    displayWord = _dictionary.ConvertSuggestionToScript(displayWord, Models.ScriptType.Latin);
                else if (inputScript == Models.ScriptType.Cyrillic)
                    displayWord = _dictionary.ConvertSuggestionToScript(displayWord, Models.ScriptType.Cyrillic);

                displayWord = PreserveCase(word, displayWord);
                if (string.IsNullOrWhiteSpace(displayWord)) continue;
                if (!seenDisplay.Add(displayWord)) continue;

                output.Add(new Suggestion(displayWord, item.Score, item.Distance));
                if (output.Count >= maxResults) break;
            }

            return output;
        }

        private int GetAdaptiveMaxDistance(string word)
        {
            if (string.IsNullOrEmpty(word)) return 2;
            if (word.Length <= 4) return 1;
            if (word.Length <= 8) return 2;
            return 3;
        }

        private List<(string Word, int Distance)> GenerateDirectVariantCandidates(string searchWord, int maxDistance)
        {
            var results = new List<(string Word, int Distance)>();
            if (string.IsNullOrEmpty(searchWord)) return results;

            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Common typo pattern: adjacent transposition (e.g., "kitob" -> "ktiob")
            var chars = searchWord.ToCharArray();
            for (int i = 0; i < chars.Length - 1; i++)
            {
                if (chars[i] == chars[i + 1]) continue;
                char a = chars[i];
                chars[i] = chars[i + 1];
                chars[i + 1] = a;
                var variant = new string(chars);
                chars[i + 1] = chars[i];
                chars[i] = a;

                if (variant.Equals(searchWord, StringComparison.OrdinalIgnoreCase)) continue;
                if (!set.Add(variant)) continue;
                if (_dictionary.Contains(variant))
                    results.Add((variant, 1));
            }

            // Apostrophe normalization variants (Latin Uzbek often mixes apostrophe chars)
            var apostrophes = new[] { '\'', '\u02BB', '\u02BC', '\u2018', '\u2019' };
            foreach (var apo in apostrophes)
            {
                var variant = searchWord
                    .Replace('\'', apo)
                    .Replace('\u02BB', apo)
                    .Replace('\u02BC', apo)
                    .Replace('\u2018', apo)
                    .Replace('\u2019', apo);

                if (variant.Equals(searchWord, StringComparison.OrdinalIgnoreCase)) continue;
                if (!set.Add(variant)) continue;
                if (!_dictionary.Contains(variant)) continue;

                int dist = TextHelper.BoundedEditDistance(searchWord, variant, maxDistance);
                if (dist > 0 && dist <= maxDistance)
                    results.Add((variant, dist));
            }

            return results;
        }

        private double ScoreCandidate(string input, string candidate, int distance, int maxDistance)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(candidate))
                return 0.0;

            int inputLen = input.Length;
            int candidateLen = candidate.Length;
            int maxLen = Math.Max(inputLen, candidateLen);
            int minLen = Math.Max(1, Math.Min(inputLen, candidateLen));

            double distanceScore = 1.0 - (double)distance / (maxLen + 1);
            double prefixScore = (double)CommonPrefixLength(input, candidate) / minLen;
            double suffixScore = (double)CommonSuffixLength(input, candidate) / minLen;
            double bigramScore = DiceCoefficient(input, candidate);
            double edgeScore = (input[0] == candidate[0] ? 0.6 : 0.0) +
                               (input[inputLen - 1] == candidate[candidateLen - 1] ? 0.4 : 0.0);

            double lengthPenalty = (double)Math.Abs(inputLen - candidateLen) / maxLen;
            double distancePenalty = maxDistance > 0 ? (double)distance / maxDistance : 0.0;
            double phoneticBoost = TextHelper.IsPhoneticallySimilar(input, candidate) ? 0.08 : 0.0;
            double transposeBoost = IsSingleAdjacentTransposition(input, candidate) ? 0.08 : 0.0;

            double score =
                (0.38 * distanceScore) +
                (0.20 * bigramScore) +
                (0.15 * prefixScore) +
                (0.10 * suffixScore) +
                (0.10 * edgeScore) -
                (0.08 * lengthPenalty) -
                (0.05 * distancePenalty) +
                phoneticBoost +
                transposeBoost;

            if (distance == 1 && input[0] == candidate[0])
                score += 0.04;

            if (score < 0.0) return 0.0;
            if (score > 1.0) return 1.0;
            return score;
        }

        private int CommonPrefixLength(string a, string b)
        {
            int n = Math.Min(a.Length, b.Length);
            int i = 0;
            while (i < n && a[i] == b[i]) i++;
            return i;
        }

        private int CommonSuffixLength(string a, string b)
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

        private double DiceCoefficient(string a, string b)
        {
            if (a.Length < 2 || b.Length < 2)
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase) ? 1.0 : 0.0;

            var aBigrams = BuildBigramMultiset(a);
            var bBigrams = BuildBigramMultiset(b);
            int intersection = 0;

            foreach (var kv in aBigrams)
            {
                if (bBigrams.TryGetValue(kv.Key, out int countB))
                    intersection += Math.Min(kv.Value, countB);
            }

            int total = aBigrams.Values.Sum() + bBigrams.Values.Sum();
            if (total == 0) return 0.0;
            return (2.0 * intersection) / total;
        }

        private Dictionary<string, int> BuildBigramMultiset(string s)
        {
            var map = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < s.Length - 1; i++)
            {
                string bg = s.Substring(i, 2);
                if (map.TryGetValue(bg, out int count))
                    map[bg] = count + 1;
                else
                    map[bg] = 1;
            }
            return map;
        }

        private bool IsSingleAdjacentTransposition(string a, string b)
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

        // =====================================================================
        //  DOCUMENT CHECK  (Optimized: batch extraction + deferred suggestions)
        // =====================================================================

        /// <inheritdoc/>
        /// <remarks>
        /// Performance-optimized path:
        /// 1. Extracts full text in ONE COM call (instead of per-word interop).
        /// 2. Tokenizes in pure C# — zero COM overhead.
        /// 3. Caches dictionary hit/miss per unique word (avoids duplicate lookups).
        /// 4. Defers suggestion computation until user actually views an error.
        /// 5. Counts paragraph indices in-memory from \r characters.
        /// 6. Suspends Word screen-updating during the entire check.
        /// </remarks>
        public List<ErrorEntry> Check(Word.Range range)
        {
            var errors = new List<ErrorEntry>();
            if (range == null) return errors;

            try
            {
                var app = range.Application;
                var doc = range.Document;

                // ── 1. Suspend rendering ────────────────────────────────────
                bool wasUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;

                try
                {
                    // ── 2. One COM call for full text ───────────────────────
                    string fullText = range.Text;
                    if (string.IsNullOrEmpty(fullText))
                    {
                        Logger.Info("Текширув: ҳужжатда матн топилмади");
                        return errors;
                    }

                    int rangeStart = range.Start;

                    // ── 3. Pure C# tokenization ─────────────────────────────
                    var tokens = TextHelper.Tokenize(fullText);

                    // ── 4. Batch dictionary check with deduplication ────────
                    var knownCorrect = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var knownMisspelled = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var errorTokens = new List<WordToken>();

                    foreach (var token in tokens)
                    {
                        if (token.Normalized.Length <= 1) continue;
                        if (TextHelper.ShouldSkipWord(token.Normalized)) continue;

                        // Already confirmed correct → skip instantly
                        if (knownCorrect.Contains(token.Normalized)) continue;

                        // Already confirmed misspelled → record another occurrence
                        if (knownMisspelled.Contains(token.Normalized))
                        {
                            errorTokens.Add(token);
                            continue;
                        }

                        // First encounter — single HashSet lookup
                        if (_dictionary.Contains(token.Normalized))
                        {
                            knownCorrect.Add(token.Normalized);
                        }
                        else
                        {
                            knownMisspelled.Add(token.Normalized);
                            errorTokens.Add(token);
                        }
                    }

                    // ── 5. Build errors with in-memory paragraph tracking ───
                    int currentPara = 1;
                    int lastSearchPos = 0;

                    foreach (var token in errorTokens)
                    {
                        // Count \r between last position and this token
                        for (int p = lastSearchPos; p < token.StartIndex && p < fullText.Length; p++)
                        {
                            if (fullText[p] == '\r') currentPara++;
                        }
                        lastSearchPos = token.StartIndex;

                        // Create Word range (only COM call per error, not per word)
                        Word.Range wordRange = null;
                        try
                        {
                            int wordStart = rangeStart + token.StartIndex;
                            int wordEnd = rangeStart + token.EndIndex;
                            wordRange = doc.Range(wordStart, wordEnd);
                        }
                        catch { continue; }

                        // Context from in-memory text (zero COM calls)
                        string context = GetContextFromText(fullText, token.StartIndex, token.EndIndex);

                        var error = new ErrorEntry
                        {
                            Word = token.Original.Trim(),
                            StartIndex = wordRange.Start,
                            EndIndex = wordRange.End,
                            ParagraphIndex = currentPara,
                            Context = context,
                            Range = wordRange,
                            Severity = ErrorSeverity.Spelling,
                            DetectedAt = DateTime.Now,
                            // DEFERRED — suggestions computed on-demand when user views this error
                            SuggestionsProvider = w => GetSuggestions(w, 5)
                        };

                        errors.Add(error);
                    }

                    Logger.Info($"Текширув натижаси: {errors.Count} та хато / {tokens.Count} та сўз " +
                                $"({knownCorrect.Count} тўғри, {knownMisspelled.Count} хато сўз)");
                }
                finally
                {
                    app.ScreenUpdating = wasUpdating;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Текширув пайтида хато", ex);
            }

            return errors;
        }

        /// <summary>
        /// Extracts a context snippet from the in-memory text around a word position.
        /// Zero COM calls — pure string operations.
        /// </summary>
        private static string GetContextFromText(string text, int wordStart, int wordEnd, int contextChars = 30)
        {
            if (string.IsNullOrEmpty(text)) return "";

            // Find paragraph boundaries (\r separators)
            int paraStart = 0;
            for (int i = Math.Min(wordStart, text.Length) - 1; i >= 0; i--)
            {
                if (text[i] == '\r') { paraStart = i + 1; break; }
            }
            int paraEnd = text.IndexOf('\r', Math.Min(wordEnd, text.Length));
            if (paraEnd < 0) paraEnd = text.Length;

            int start = Math.Max(paraStart, wordStart - contextChars);
            int end = Math.Min(paraEnd, wordEnd + contextChars);

            var result = new System.Text.StringBuilder();
            if (start > paraStart) result.Append("...");
            result.Append(text.Substring(start, end - start));
            if (end < paraEnd) result.Append("...");

            return result.ToString().Trim();
        }

        // =====================================================================
        //  HIGHLIGHTING
        // =====================================================================

        /// <inheritdoc/>
        public void HighlightErrors(List<ErrorEntry> errors)
        {
            if (errors == null) return;
            foreach (var error in errors)
                if (error != null && !error.IsResolved)
                    DocumentHighlightService.Highlight(error.Range, Word.WdColor.wdColorRed);
        }

        /// <inheritdoc/>
        /// <remarks>Only clears temporary marks owned by this add-in session.</remarks>
        public void ClearHighlights(Word.Document document)
        {
            DocumentHighlightService.ClearDocument(document);
        }

        // =====================================================================
        //  HELPERS
        // =====================================================================

        /// <summary>
        /// Preserves the original case pattern of the misspelled word in the suggestion.
        /// E.g., if original is "Маслахат" (title case), suggestion becomes "Маслаҳат".
        /// </summary>
        private string PreserveCase(string original, string suggestion)
        {
            if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(suggestion))
                return suggestion;

            original = original.Trim();

            // All uppercase
            if (original.All(c => !char.IsLetter(c) || char.IsUpper(c)))
                return suggestion.ToUpperInvariant();

            // Title case (first letter uppercase)
            if (char.IsUpper(original[0]) && original.Skip(1).All(c => !char.IsLetter(c) || char.IsLower(c)))
            {
                if (suggestion.Length > 0)
                    return char.ToUpperInvariant(suggestion[0]) + suggestion.Substring(1);
            }

            return suggestion;
        }
    }
}
