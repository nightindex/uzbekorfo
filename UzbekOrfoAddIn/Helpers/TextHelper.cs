using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Static utility class for Uzbek text processing — normalization, tokenization,
    /// edit distance, and phonetic similarity.
    /// </summary>
    public static class TextHelper
    {
        // Uzbek Cyrillic alphabet characters (lowercase)
        private static readonly HashSet<char> UzbekCyrillicChars = new HashSet<char>(
            "абвгғдеёжзийклмнопрстуўфхҳцчшъэюя".ToCharArray());

        // Uzbek Latin alphabet characters (lowercase)
        private static readonly HashSet<char> UzbekLatinChars = new HashSet<char>(
            "abcdefghijklmnopqrstuvwxyz'ʻʼ".ToCharArray());

        // Characters that are considered word characters in Uzbek
        private static readonly Regex WordPattern = new Regex(
            @"[\p{L}ʻʼ'']+", RegexOptions.Compiled);

        // Common Uzbek confusable pairs (for phonetic similarity)
        private static readonly Dictionary<char, char> PhoneticMap = new Dictionary<char, char>
        {
            {'х', 'ҳ'}, {'ҳ', 'х'},
            {'к', 'қ'}, {'қ', 'к'},
            {'г', 'ғ'}, {'ғ', 'г'},
            {'у', 'ў'}, {'ў', 'у'},
            {'о', 'ў'}, {'е', 'э'},
            {'э', 'е'}, {'и', 'й'},
            {'й', 'и'}
        };

        // =====================================================================
        //  NORMALIZATION
        // =====================================================================

        /// <summary>
        /// Normalizes a word for dictionary lookup: lowercase, trim, remove punctuation edges.
        /// </summary>
        public static string NormalizeWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return string.Empty;

            word = word.Trim().ToLowerInvariant();

            // Strip ALL leading non-letter / non-apostrophe characters
            int start = 0;
            while (start < word.Length && !IsWordChar(word[start]))
                start++;

            // Strip ALL trailing non-letter / non-apostrophe characters
            int end = word.Length - 1;
            while (end >= start && !IsWordChar(word[end]))
                end--;

            if (start > end) return string.Empty;

            word = word.Substring(start, end - start + 1);

            return word;
        }

        /// <summary>
        /// Returns true if the character is valid at the edge of an Uzbek word
        /// (letter or apostrophe variant used in oʻzbek / gʻ).
        /// </summary>
        private static bool IsWordChar(char c)
        {
            if (char.IsLetter(c)) return true;
            // Uzbek apostrophe variants: ' (U+0027), ʻ (U+02BB), ʼ (U+02BC),
            // curly quotes ' (U+2018) ' (U+2019)
            return c == '\'' || c == 'ʻ' || c == 'ʼ' || c == '\u2018' || c == '\u2019';
        }

        // =====================================================================
        //  TOKENIZATION
        // =====================================================================

        /// <summary>
        /// Splits text into individual words (tokens), preserving their position info.
        /// </summary>
        public static List<WordToken> Tokenize(string text)
        {
            var tokens = new List<WordToken>();
            if (string.IsNullOrEmpty(text)) return tokens;

            var matches = WordPattern.Matches(text);
            foreach (Match match in matches)
            {
                var normalized = NormalizeWord(match.Value);
                if (normalized.Length >= 1)
                {
                    tokens.Add(new WordToken
                    {
                        Original = match.Value,
                        Normalized = normalized,
                        StartIndex = match.Index,
                        EndIndex = match.Index + match.Length
                    });
                }
            }

            return tokens;
        }

        // =====================================================================
        //  EDIT DISTANCE (Damerau-Levenshtein)
        // =====================================================================

        /// <summary>
        /// Computes the Damerau-Levenshtein edit distance between two strings.
        /// Supports insertions, deletions, substitutions, and transpositions.
        /// </summary>
        public static int EditDistance(string source, string target)
        {
            if (string.IsNullOrEmpty(source)) return target?.Length ?? 0;
            if (string.IsNullOrEmpty(target)) return source.Length;

            int n = source.Length;
            int m = target.Length;

            // Quick check for max distance
            if (Math.Abs(n - m) > 3) return Math.Abs(n - m);

            var d = new int[n + 1, m + 1];

            for (int i = 0; i <= n; i++) d[i, 0] = i;
            for (int j = 0; j <= m; j++) d[0, j] = j;

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = source[i - 1] == target[j - 1] ? 0 : 1;

                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);

                    // Transposition
                    if (i > 1 && j > 1 &&
                        source[i - 1] == target[j - 2] &&
                        source[i - 2] == target[j - 1])
                    {
                        d[i, j] = Math.Min(d[i, j], d[i - 2, j - 2] + cost);
                    }
                }
            }

            return d[n, m];
        }

        /// <summary>
        /// Bounded Damerau-Levenshtein: returns early (maxDistance+1) when the
        /// minimum possible distance exceeds the threshold. Uses rolling 3-row
        /// arrays instead of a full n×m matrix — ∼70% faster for large words.
        /// </summary>
        public static int BoundedEditDistance(string source, string target, int maxDistance)
        {
            if (string.IsNullOrEmpty(source)) return target?.Length ?? 0;
            if (string.IsNullOrEmpty(target)) return source.Length;

            int n = source.Length;
            int m = target.Length;
            int over = maxDistance + 1;

            if (Math.Abs(n - m) > maxDistance) return over;

            // Three rolling rows: prev2, prev, curr
            var prev2 = new int[m + 1];
            var prev  = new int[m + 1];
            var curr  = new int[m + 1];

            for (int j = 0; j <= m; j++) prev[j] = j;

            for (int i = 1; i <= n; i++)
            {
                curr[0] = i;
                int rowMin = curr[0];

                for (int j = 1; j <= m; j++)
                {
                    int cost = source[i - 1] == target[j - 1] ? 0 : 1;

                    curr[j] = Math.Min(
                        Math.Min(curr[j - 1] + 1, prev[j] + 1),
                        prev[j - 1] + cost);

                    // Transposition
                    if (i > 1 && j > 1 &&
                        source[i - 1] == target[j - 2] &&
                        source[i - 2] == target[j - 1])
                    {
                        curr[j] = Math.Min(curr[j], prev2[j - 2] + cost);
                    }

                    if (curr[j] < rowMin) rowMin = curr[j];
                }

                // Early termination: if the best value in this row already
                // exceeds the threshold, no subsequent row can bring it down.
                if (rowMin > maxDistance) return over;

                // Rotate rows
                var tmp = prev2;
                prev2 = prev;
                prev = curr;
                curr = tmp;
            }

            return prev[m];
        }

        // =====================================================================
        //  PHONETIC SIMILARITY
        // =====================================================================

        /// <summary>
        /// Checks if two words are phonetically similar in Uzbek
        /// (differ only by commonly confused character pairs like х↔ҳ, к↔қ, etc.).
        /// </summary>
        public static bool IsPhoneticallySimilar(string word1, string word2)
        {
            if (word1.Length != word2.Length) return false;

            int differences = 0;
            for (int i = 0; i < word1.Length; i++)
            {
                if (word1[i] != word2[i])
                {
                    differences++;
                    if (differences > 2) return false;

                    // Check if the differing characters are a known confusable pair
                    if (!PhoneticMap.ContainsKey(word1[i]) || PhoneticMap[word1[i]] != word2[i])
                    {
                        if (!PhoneticMap.ContainsKey(word2[i]) || PhoneticMap[word2[i]] != word1[i])
                            return false;
                    }
                }
            }

            return differences > 0;
        }

        /// <summary>
        /// Generates phonetic variants of a word by swapping confusable characters.
        /// </summary>
        public static List<string> GeneratePhoneticVariants(string word)
        {
            var variants = new HashSet<string>();
            var chars = word.ToCharArray();

            for (int i = 0; i < chars.Length; i++)
            {
                if (PhoneticMap.ContainsKey(chars[i]))
                {
                    var original = chars[i];
                    chars[i] = PhoneticMap[original];
                    variants.Add(new string(chars));
                    chars[i] = original;
                }
            }

            return variants.ToList();
        }

        // =====================================================================
        //  SCRIPT DETECTION
        // =====================================================================

        /// <summary>
        /// Returns true if the character is a Cyrillic letter.
        /// </summary>
        public static bool IsCyrillic(char c) =>
            (c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') ||
            c == 'ё' || c == 'Ё' || c == 'ў' || c == 'Ў' ||
            c == 'ғ' || c == 'Ғ' || c == 'ҳ' || c == 'Ҳ' ||
            c == 'қ' || c == 'Қ';

        /// <summary>
        /// Returns true if the character is a Latin letter.
        /// </summary>
        public static bool IsLatin(char c) =>
            (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');

        /// <summary>
        /// Returns true if the text is predominantly Cyrillic.
        /// </summary>
        public static bool IsPredominantlyCyrillic(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            int cyrillic = text.Count(IsCyrillic);
            int latin = text.Count(IsLatin);
            return cyrillic > latin;
        }

        /// <summary>
        /// Returns true if a word looks like a number, abbreviation, or non-linguistic token.
        /// </summary>
        public static bool ShouldSkipWord(string word)
        {
            if (string.IsNullOrEmpty(word)) return true;
            if (word.Length <= 1) return true;
            if (word.All(char.IsDigit)) return true;
            if (word.All(c => char.IsUpper(c) || c == '.')) return true; // abbreviations like "BMT"

            // Reject words containing brackets, braces, or other non-word characters
            // Valid interior characters: letters, apostrophe variants, hyphen, soft-hyphen
            for (int i = 0; i < word.Length; i++)
            {
                char c = word[i];
                if (char.IsLetter(c)) continue;
                if (c == '-' || c == '\u00AD') continue;               // hyphen, soft-hyphen
                if (c == '\'' || c == 'ʻ' || c == 'ʼ') continue;       // apostrophes
                if (c == '\u2018' || c == '\u2019') continue;          // curly quotes
                return true; // brackets, digits, symbols, etc. → skip
            }

            return false;
        }
    }

    /// <summary>
    /// Represents a single word token extracted from text, with position info.
    /// </summary>
    public class WordToken
    {
        public string Original { get; set; }
        public string Normalized { get; set; }
        public int StartIndex { get; set; }
        public int EndIndex { get; set; }
    }
}
