using System;
using System.Collections.Generic;
using UzbekOrfoAddIn.Helpers;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Normalizes Uzbek apostrophe/tutuq symbols in text and Word ranges.
    /// </summary>
    public sealed class UzbekApostropheService
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

        public string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            var chars = text.ToCharArray();
            bool changed = false;

            for (int i = 0; i < chars.Length; i++)
            {
                if (!ApostropheLikeChars.Contains(chars[i]))
                    continue;

                char replacement;
                if (!TryGetCanonicalUzbekApostrophe(text, i, out replacement))
                    continue;

                if (chars[i] != replacement)
                {
                    chars[i] = replacement;
                    changed = true;
                }
            }

            return changed ? new string(chars) : text;
        }

        public int NormalizeRange(Word.Range range, string undoLabel)
        {
            if (range == null) return 0;

            string text = range.Text ?? string.Empty;
            if (text.Length == 0) return 0;

            var edits = new List<Tuple<int, char>>();
            for (int i = 0; i < text.Length; i++)
            {
                char current = text[i];
                if (!ApostropheLikeChars.Contains(current))
                    continue;

                char replacement;
                if (!TryGetCanonicalUzbekApostrophe(text, i, out replacement))
                    continue;

                if (current != replacement)
                    edits.Add(Tuple.Create(i, replacement));
            }

            if (edits.Count == 0) return 0;

            DocumentHelper.BeginUndoRecord(undoLabel);
            try
            {
                for (int i = edits.Count - 1; i >= 0; i--)
                {
                    int offset = edits[i].Item1;
                    char replacement = edits[i].Item2;
                    var oneChar = range.Document.Range(range.Start + offset, range.Start + offset + 1);
                    oneChar.Text = replacement.ToString();
                }
            }
            finally
            {
                DocumentHelper.EndUndoRecord();
            }

            return edits.Count;
        }

        private bool TryGetCanonicalUzbekApostrophe(string text, int index, out char replacement)
        {
            replacement = '\0';

            char left = index > 0 ? text[index - 1] : '\0';
            char right = index + 1 < text.Length ? text[index + 1] : '\0';

            bool leftLatin = IsLatinLetter(left);
            bool rightLatin = IsLatinLetter(right);

            if (!leftLatin || !rightLatin)
                return false;

            if (IsOGLetter(left))
            {
                replacement = '‘';
                return true;
            }

            replacement = '’';
            return true;
        }

        private static bool IsOGLetter(char c)
        {
            return c == 'O' || c == 'o' || c == 'G' || c == 'g';
        }

        private static bool IsLatinLetter(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }
    }
}
