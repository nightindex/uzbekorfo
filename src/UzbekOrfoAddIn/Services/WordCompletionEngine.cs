using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>Immutable, host-independent prefix index. Never owns a second dictionary.</summary>
    public sealed class WordCompletionEngine
    {
        private readonly string[] _words;
        private readonly string[] _endings;
        private readonly Func<string, bool> _validForm;

        public WordCompletionEngine(IEnumerable<string> words, IEnumerable<string> endings = null,
            Func<string, bool> validForm = null, CancellationToken cancellation = default(CancellationToken))
        {
            _words = Index(words, cancellation);
            _endings = Index(endings ?? Enumerable.Empty<string>(), cancellation);
            _validForm = validForm;
        }

        private static string[] Index(IEnumerable<string> values, CancellationToken cancellation) => values
            .Select(value => { cancellation.ThrowIfCancellationRequested(); return value; })
            .Where(IsToken).Select(TextHelper.NormalizeWord).Where(w => w.Length > 0)
            .Distinct(StringComparer.Ordinal).OrderBy(w => w, StringComparer.Ordinal).ToArray();

        public static bool IsToken(string word)
        {
            if (string.IsNullOrEmpty(word) || word.Length > 80 || !char.IsLetter(word[0])) return false;
            bool latin = false, cyrillic = false;
            foreach (char c in word)
            {
                if (!IsTokenCharacter(c)) return false;
                latin |= (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
                cyrillic |= c >= '\u0400' && c <= '\u04ff';
            }
            return latin != cyrillic; // Reject mixed scripts and non-Uzbek-only tokens.
        }

        public static bool IsTokenCharacter(char c) => char.IsLetter(c) || c == '-' ||
            c == '\'' || c == '\u2018' || c == '\u2019' || c == '\u02bb' || c == '\u02bc' || c == '`';

        private static int LowerBound(string[] values, string prefix)
        {
            int low = 0, high = values.Length;
            while (low < high)
            {
                int middle = low + (high - low) / 2;
                if (string.CompareOrdinal(values[middle], prefix) < 0) low = middle + 1;
                else high = middle;
            }
            return low;
        }

        /// <param name="deferMorphologyValidation">Generation-only mode for the Word adapter.
        /// Caller MUST validate every returned candidate on the owning thread before display.</param>
        public string[] Complete(string prefix, int count = 3, CancellationToken cancellation = default(CancellationToken),
            bool deferMorphologyValidation = false, IDictionary<string, int> acceptanceCounts = null)
        {
            cancellation.ThrowIfCancellationRequested();
            if (count <= 0 || !IsToken(prefix)) return new string[0];
            count = Math.Min(count, 10);
            string normalized = TextHelper.NormalizeWord(prefix);
            if (normalized.Length < 2) return new string[0];
            var found = new HashSet<string>(StringComparer.Ordinal);
            // Small opt-in overlay; learned completions outside the lexical retrieval window
            // must still be proposed. The Word adapter validates them before display.
            if (acceptanceCounts != null)
                foreach (var word in acceptanceCounts.Keys.Take(1000))
                {
                    cancellation.ThrowIfCancellationRequested();
                    if (IsToken(word) && word.Length > normalized.Length && word.StartsWith(normalized, StringComparison.Ordinal))
                        found.Add(word);
                }
            int start = LowerBound(_words, normalized);
            // Fixed-size lexical candidate window. Never enumerate the entire prefix range.
            for (int i = start; i < _words.Length && i < start + 96; i++)
            {
                cancellation.ThrowIfCancellationRequested();
                if (!_words[i].StartsWith(normalized, StringComparison.Ordinal)) break;
                if (_words[i].Length > normalized.Length) found.Add(_words[i]);
            }
            // Conservative partial-ending completion only. No recursive suffix expansion.
            // Every generated form must pass the existing analyzer, including root flags.
            if (_validForm != null || deferMorphologyValidation)
            {
                var budget = Stopwatch.StartNew();
                int checkedForms = 0;
                for (int split = normalized.Length - 1; split >= 2 && split >= normalized.Length - 16; split--)
                {
                    string stem = normalized.Substring(0, split), tail = normalized.Substring(split);
                    int root = LowerBound(_words, stem);
                    if (root == _words.Length || _words[root] != stem) continue;
                    int first = LowerBound(_endings, tail);
                    for (int j = first; j < _endings.Length && _endings[j].StartsWith(tail, StringComparison.Ordinal); j++)
                    {
                        cancellation.ThrowIfCancellationRequested();
                        if (++checkedForms > 24 || budget.ElapsedMilliseconds >= 8) break;
                        string form = stem + _endings[j];
                        if (form.Length > normalized.Length && !found.Contains(form) &&
                            (deferMorphologyValidation || _validForm(form))) found.Add(form);
                    }
                    if (checkedForms >= 24 || budget.ElapsedMilliseconds >= 8) break;
                }
            }
            // No fabricated frequency model: deterministic short-completion ranking for MVP.
            return found.OrderByDescending(w => acceptanceCounts != null && acceptanceCounts.ContainsKey(w) ? acceptanceCounts[w] : 0)
                .ThenBy(w => w.Length).ThenBy(w => w, StringComparer.Ordinal).Take(count)
                .Select(w => PreservePrefix(prefix, w.Substring(normalized.Length))).ToArray();
        }

        private static string PreservePrefix(string prefix, string tail)
        {
            bool uppercase = prefix.Where(char.IsLetter).All(char.IsUpper);
            // Keep the exact typed casing and apostrophes; insertion only adds this tail.
            return prefix + (uppercase ? tail.ToUpperInvariant() : tail);
        }

        public static string[] ValidateCandidates(IEnumerable<string> proposals, Func<string, bool> isValid, int count)
        {
            var accepted = new List<string>();
            if (count <= 0) return accepted.ToArray();
            var budget = Stopwatch.StartNew();
            foreach (var proposal in proposals.Take(10))
            {
                if (isValid(proposal)) accepted.Add(proposal);
                if (accepted.Count >= count || budget.ElapsedMilliseconds >= 5) break;
            }
            return accepted.ToArray();
        }
    }
}
