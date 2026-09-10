using System;
using System.Collections.Generic;
using System.Text;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Sorts Uzbek words correctly according to the official Uzbek alphabet order
    /// for both Cyrillic and Latin scripts.
    /// 
    /// Cyrillic: А Б В Г Ғ Д Е Ё Ж З И Й К Қ Л М Н О П Р С Т У Ў Ф Х Ҳ Ц Ч Ш Ъ Э Ю Я
    /// Latin:    A B Ch D E F G G' H I J K L M N Ng O O' P Q R S Sh T U V X Y Z
    ///
    /// Mixed script words are grouped: all Cyrillic words first, then all Latin words.
    /// Falls back to ordinal comparison for characters outside the Uzbek alphabet.
    /// </summary>
    public sealed class UzbekStringComparer : IComparer<string>
    {
        public static readonly UzbekStringComparer Instance = new UzbekStringComparer();
        private UzbekStringComparer() { }

        // ─── Uzbek Cyrillic alphabet → sort weight (lowercase key) ───────────
        private static readonly Dictionary<char, int> _cyrWeights = new Dictionary<char, int>
        {
            { 'а',  1 }, { 'б',  2 }, { 'в',  3 }, { 'г',  4 }, { 'ғ',  5 },
            { 'д',  6 }, { 'е',  7 }, { 'ё',  8 }, { 'ж',  9 }, { 'з', 10 },
            { 'и', 11 }, { 'й', 12 }, { 'к', 13 }, { 'қ', 14 }, { 'л', 15 },
            { 'м', 16 }, { 'н', 17 }, { 'о', 18 }, { 'п', 19 }, { 'р', 20 },
            { 'с', 21 }, { 'т', 22 }, { 'у', 23 }, { 'ў', 24 }, { 'ф', 25 },
            { 'х', 26 }, { 'ҳ', 27 }, { 'ц', 28 }, { 'ч', 29 }, { 'ш', 30 },
            { 'ъ', 31 }, { 'э', 32 }, { 'ю', 33 }, { 'я', 34 }
        };

        // ─── Uzbek Latin tokens (digraphs listed longest-first per first char) ─
        // Each entry: (token_lowercase, sort_weight)
        // Official order: A(1) B(2) Ch(3) D(4) E(5) F(6) G(7) G'(8) H(9) I(10)
        //                 J(11) K(12) L(13) M(14) N(15) Ng(16) O(17) O'(18) P(19)
        //                 Q(20) R(21) S(22) Sh(23) T(24) U(25) V(26) X(27) Y(28) Z(29)
        private static readonly (string Token, int Weight)[] _latTokens =
        {
            ("a",   1), ("b",   2), ("ch",  3), ("d",   4), ("e",   5),
            ("f",   6), ("g'",  8), ("g",   7), ("h",   9), ("i",  10),
            ("j",  11), ("k",  12), ("l",  13), ("m",  14), ("ng", 16),
            ("n",  15), ("o'", 18), ("o",  17), ("p",  19), ("q",  20),
            ("r",  21), ("s",  22), ("sh", 23), ("t",  24), ("u",  25),
            ("v",  26), ("x",  27), ("y",  28), ("z",  29)
        };

        // Lookup: first char → tokens sorted longest-first (so digraphs win over singles)
        private static readonly Dictionary<char, (string Token, int Weight)[]> _latLookup;

        static UzbekStringComparer()
        {
            var temp = new Dictionary<char, List<(string, int)>>();
            foreach (var t in _latTokens)
            {
                char first = t.Token[0];
                if (!temp.ContainsKey(first))
                    temp[first] = new List<(string, int)>();
                temp[first].Add((t.Token, t.Weight));
            }

            _latLookup = new Dictionary<char, (string, int)[]>();
            foreach (var kv in temp)
            {
                var list = kv.Value;
                // Sort longest token first so "ch"/"ng"/"sh"/"g'"/"o'" beat single letters
                list.Sort((a, b) => b.Item1.Length.CompareTo(a.Item1.Length));
                _latLookup[kv.Key] = list.ToArray();
            }
        }

        // ─── Public interface ────────────────────────────────────────────────

        public int Compare(string x, string y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            bool xCyr = IsCyrillicWord(x);
            bool yCyr = IsCyrillicWord(y);

            // Group scripts: all Cyrillic words first, then all Latin
            if (xCyr != yCyr) return xCyr ? -1 : 1;

            return xCyr ? CompareCyrillic(x, y) : CompareLatin(x, y);
        }

        // ─── Cyrillic comparison ─────────────────────────────────────────────

        private static int CompareCyrillic(string x, string y)
        {
            string xl = x.ToLowerInvariant();
            string yl = y.ToLowerInvariant();
            int len = Math.Min(xl.Length, yl.Length);

            for (int i = 0; i < len; i++)
            {
                int wx = _cyrWeights.TryGetValue(xl[i], out int vx) ? vx : (int)xl[i] + 1000;
                int wy = _cyrWeights.TryGetValue(yl[i], out int vy) ? vy : (int)yl[i] + 1000;
                if (wx != wy) return wx.CompareTo(wy);
            }

            return xl.Length.CompareTo(yl.Length);
        }

        // ─── Latin comparison (digraph-aware) ───────────────────────────────

        private static int CompareLatin(string x, string y)
        {
            string xl = NormalizeApostrophes(x.ToLowerInvariant());
            string yl = NormalizeApostrophes(y.ToLowerInvariant());

            int xi = 0, yi = 0;
            while (xi < xl.Length && yi < yl.Length)
            {
                NextLatinToken(xl, xi, out int wx, out int cx);
                xi += cx;

                NextLatinToken(yl, yi, out int wy, out int cy);
                yi += cy;

                if (wx != wy) return wx.CompareTo(wy);
            }

            bool xDone = xi >= xl.Length;
            bool yDone = yi >= yl.Length;
            if (xDone && yDone) return 0;
            return xDone ? -1 : 1;  // shorter word comes first
        }

        /// <summary>
        /// Reads the next logical Uzbek Latin "letter" (1 or 2 chars) starting at
        /// position <paramref name="i"/> in <paramref name="s"/> and returns its
        /// sort weight and how many characters were consumed.
        /// </summary>
        private static void NextLatinToken(string s, int i, out int weight, out int consumed)
        {
            char c = s[i];
            if (_latLookup.TryGetValue(c, out var tokens))
            {
                // Tokens are already sorted longest-first; pick first match
                foreach (var (token, w) in tokens)
                {
                    if (i + token.Length <= s.Length &&
                        string.Compare(s, i, token, 0, token.Length, StringComparison.Ordinal) == 0)
                    {
                        weight = w;
                        consumed = token.Length;
                        return;
                    }
                }
            }

            // Character not in Uzbek Latin alphabet — sort after all known letters
            weight = (int)c + 10000;
            consumed = 1;
        }

        // ─── Helpers ─────────────────────────────────────────────────────────

        private static bool IsCyrillicWord(string w)
        {
            int cyr = 0, lat = 0;
            foreach (char c in w)
            {
                char lo = char.ToLowerInvariant(c);
                if ((lo >= 'а' && lo <= 'я') || _cyrWeights.ContainsKey(lo)) cyr++;
                else if (lo >= 'a' && lo <= 'z') lat++;
            }
            return cyr >= lat;  // treat ties as Cyrillic
        }

        private static bool IsApostrophe(char c) =>
            c == '\'' || c == '\u2018' || c == '\u2019' ||
            c == '\u02BC' || c == '\u02BB' || c == '`';

        private static string NormalizeApostrophes(string s)
        {
            for (int i = 0; i < s.Length; i++)
            {
                if (IsApostrophe(s[i]) && s[i] != '\'')
                {
                    var sb = new StringBuilder(s);
                    for (int j = i; j < sb.Length; j++)
                        if (IsApostrophe(sb[j])) sb[j] = '\'';
                    return sb.ToString();
                }
            }
            return s;
        }
    }
}
