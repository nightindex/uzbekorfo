using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using UzbekOrfoAddIn.Helpers;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Owns temporary marks on the Word UI thread. Existing or mixed underline
    /// formatting is never overwritten. Only ranges marked by this session are cleared.
    /// </summary>
    internal static class DocumentHighlightService
    {
        private sealed class Mark
        {
            public Word.Range Range;
            public Word.WdColor OriginalColor;
            public Word.WdColor AppliedColor;
        }

        private static readonly List<Mark> Marks = new List<Mark>();

        public static void Highlight(Word.Range range, Word.WdColor color)
        {
            if (range == null) return;
            Word.Range owned = null;
            try
            {
                // Mixed formatting is wdUndefined. Preserve it, along with any
                // deliberate underline, rather than flattening the selection.
                if (range.Underline != Word.WdUnderline.wdUnderlineNone ||
                    (int)range.Font.UnderlineColor == (int)Word.WdConstants.wdUndefined)
                    return;
                owned = range.Duplicate;
                var mark = new Mark { Range = owned, OriginalColor = owned.Font.UnderlineColor, AppliedColor = color };
                Marks.Add(mark); // Keep restoration information even if applying a mark fails halfway.
                owned = null;
                mark.Range.Underline = Word.WdUnderline.wdUnderlineWavy;
                mark.Range.Font.UnderlineColor = color;
            }
            catch (Exception ex) { Logger.Warn("Highlight skipped: " + ex.Message); }
            finally { Release(owned); }
        }

        public static void ClearDocument(Word.Document document)
        {
            if (document == null) return;
            Clear(mark => DocumentHelper.IsSameDocument(mark.Range.Document, document));
        }

        public static void ClearRange(Word.Range range)
        {
            if (range == null) return;
            Clear(mark => mark.Range.InStory(range) &&
                mark.Range.Start <= range.End && range.Start <= mark.Range.End);
        }

        public static void ClearAll() => Clear(mark => true);

        private static void Clear(Func<Mark, bool> matches)
        {
            for (int i = Marks.Count - 1; i >= 0; i--)
            {
                var mark = Marks[i];
                try
                {
                    if (!matches(mark)) continue;
                    if (mark.Range.Underline == Word.WdUnderline.wdUnderlineWavy &&
                        mark.Range.Font.UnderlineColor == mark.AppliedColor)
                    {
                        mark.Range.Underline = Word.WdUnderline.wdUnderlineNone;
                        mark.Range.Font.UnderlineColor = mark.OriginalColor;
                    }
                    else
                    {
                        // Restore individual characters so a later user formatting change
                        // in part of a marked range is preserved. Live ranges follow edits.
                        for (int c = 1; c <= mark.Range.Characters.Count; c++)
                        {
                            Word.Range character = null;
                            try
                            {
                                character = mark.Range.Characters[c];
                                if (character.Underline == Word.WdUnderline.wdUnderlineWavy &&
                                    character.Font.UnderlineColor == mark.AppliedColor)
                                {
                                    character.Underline = Word.WdUnderline.wdUnderlineNone;
                                    character.Font.UnderlineColor = mark.OriginalColor;
                                }
                            }
                            finally { Release(character); }
                        }
                    }
                }
                catch (Exception ex) { Logger.Warn("Highlight cleanup skipped: " + ex.Message); }
                Marks.RemoveAt(i);
                Release(mark.Range);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                try { Marshal.ReleaseComObject(value); } catch { }
        }
    }
}
