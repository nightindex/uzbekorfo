using System;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Static utility class wrapping common Word Interop operations.
    /// Provides safe access to the active document, selection, and text manipulation.
    /// </summary>
    public static class DocumentHelper
    {
        /// <summary>
        /// Returns the Word Application instance.
        /// </summary>
        public static Word.Application App
        {
            get
            {
                try { return Globals.ThisAddIn?.Application; }
                catch { return null; }
            }
        }

        /// <summary>
        /// Returns the currently active document, or null if none is open.
        /// </summary>
        public static Word.Document ActiveDoc
        {
            get
            {
                try { return App.ActiveDocument; }
                catch { return null; }
            }
        }

        /// <summary>
        /// Returns the current selection in the active document.
        /// </summary>
        public static Word.Selection Selection
        {
            get
            {
                try { return App?.Selection; }
                catch { return null; }
            }
        }

        /// <summary>
        /// Returns the selected text, or empty string if nothing is selected.
        /// </summary>
        public static string GetSelectedText()
        {
            var sel = Selection;
            if (sel == null || sel.Type == Word.WdSelectionType.wdNoSelection)
                return string.Empty;
            return sel.Text?.Trim() ?? string.Empty;
        }

        /// <summary>
        /// Returns the word at the cursor. If multiple characters are selected, returns
        /// the selection text. If the cursor is an insertion point (or only 1 char selected),
        /// expands to the full word under the cursor.
        /// </summary>
        public static string GetSelectedWord()
        {
            var sel = Selection;
            if (sel == null || sel.Type == Word.WdSelectionType.wdNoSelection)
                return string.Empty;

            string text = sel.Text?.Trim() ?? string.Empty;

            // If it's an insertion point or only 1 character, expand to the whole word
            if (text.Length <= 1)
            {
                try
                {
                    var wordRange = sel.Range.Words[1];
                    if (wordRange != null)
                    {
                        string wordText = wordRange.Text?.Trim() ?? string.Empty;
                        if (wordText.Length > 0)
                            return wordText;
                    }
                }
                catch { }
            }

            return text;
        }

        /// <summary>
        /// Returns the selected text if any, otherwise the entire document text.
        /// </summary>
        public static string GetSelectedOrAllText()
        {
            return GetTargetRange()?.Text ?? string.Empty;
        }

        /// <summary>
        /// Returns a Word.Range covering the selection if text is selected,
        /// or the entire document content if nothing is selected.
        /// </summary>
        public static Word.Range GetTargetRange()
        {
            return GetTargetRange(Selection, ActiveDoc);
        }

        /// <summary>Explicit host boundary for callers holding a particular Word window.</summary>
        public static Word.Range GetTargetRange(Word.Selection sel, Word.Document doc)
        {
            if (sel != null && sel.Type != Word.WdSelectionType.wdSelectionIP
                && sel.Type != Word.WdSelectionType.wdNoSelection
                && sel.Range.Start != sel.Range.End)
            {
                return sel.Range.Duplicate;
            }

            return doc?.Content;
        }

        /// <summary>
        /// Replaces text after restoring any temporary add-in marks.
        /// Word replacements exclude trailing whitespace by default. Whole-range
        /// transformations pass false because their output already contains that whitespace.
        /// </summary>
        public static void ReplaceRangeText(Word.Range range, string newText, bool preserveTrailingWhitespace = true)
        {
            if (range == null) return;
            Services.DocumentHighlightService.ClearRange(range);

            // Word cannot remove the main story's final paragraph mark. Assigning
            // output that includes it inserts another paragraph before that mark.
            // Preserve the existing final mark and replace only the preceding text.
            if (!preserveTrailingWhitespace && newText != null && newText.EndsWith("\r", StringComparison.Ordinal) &&
                range.StoryType == Word.WdStoryType.wdMainTextStory && range.End == range.Document.Content.End &&
                (range.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal))
            {
                range.End--;
                newText = newText.Substring(0, newText.Length - 1);
            }

            // Word's Words collection includes trailing space in the range.
            // Shrink the range to exclude trailing whitespace before replacing.
            try
            {
                string current = range.Text;
                if (preserveTrailingWhitespace && current != null && current.Length > 0 && char.IsWhiteSpace(current[current.Length - 1]))
                {
                    // Move End back to exclude trailing whitespace
                    int trimmedLen = current.TrimEnd().Length;
                    range.End = range.Start + trimmedLen;
                }
            }
            catch { }

            range.Text = newText;
        }

        /// <summary>Compare live COM identity, never file names (which can change on Save As).</summary>
        public static bool IsSameDocument(Word.Document left, Word.Document right)
        {
            if (left == null || right == null) return false;
            if (ReferenceEquals(left, right)) return true;
            if (!System.Runtime.InteropServices.Marshal.IsComObject(left) ||
                !System.Runtime.InteropServices.Marshal.IsComObject(right)) return false;
            IntPtr leftId = IntPtr.Zero, rightId = IntPtr.Zero;
            try
            {
                leftId = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(left);
                rightId = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(right);
                return leftId == rightId;
            }
            catch { return false; }
            finally
            {
                if (leftId != IntPtr.Zero) System.Runtime.InteropServices.Marshal.Release(leftId);
                if (rightId != IntPtr.Zero) System.Runtime.InteropServices.Marshal.Release(rightId);
            }
        }

        public static bool TryReplaceError(Models.ErrorEntry error, Word.Document document, string replacement)
        {
            if (error == null || error.IsResolved || error.Range == null || replacement == null)
                return false;
            try
            {
                var range = error.Range;
                if (!IsSameDocument(range.Document, document) ||
                    !string.Equals(range.Text, error.OriginalText ?? error.Word, StringComparison.Ordinal))
                    return false;
                ReplaceRangeText(range, replacement, preserveTrailingWhitespace: false);
                error.IsResolved = true;
                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn("Correction skipped: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Gets a context snippet around a word range (surrounding text).
        /// </summary>
        public static string GetContext(Word.Range wordRange, int contextChars = 30)
        {
            if (wordRange == null) return string.Empty;

            try
            {
                var para = wordRange.Paragraphs[1];
                var paraText = para.Range.Text ?? string.Empty;
                var wordText = wordRange.Text ?? string.Empty;

                int idx = paraText.IndexOf(wordText, StringComparison.Ordinal);
                if (idx < 0) return paraText.Length > 80 ? paraText.Substring(0, 80) + "..." : paraText;

                int start = Math.Max(0, idx - contextChars);
                int end = Math.Min(paraText.Length, idx + wordText.Length + contextChars);

                var sb = new StringBuilder();
                if (start > 0) sb.Append("...");
                sb.Append(paraText.Substring(start, end - start));
                if (end < paraText.Length) sb.Append("...");

                return sb.ToString().Trim();
            }
            catch
            {
                return wordRange.Text ?? string.Empty;
            }
        }

        /// <summary>
        /// Navigates the document to the given range and selects it.
        /// </summary>
        public static void GoToRange(Word.Range range)
        {
            if (range == null) return;
            range.Select();
            App.ActiveWindow.ScrollIntoView(range);
        }

        /// <summary>
        /// Begins a custom undo record for batching multiple operations.
        /// </summary>
        public static void BeginUndoRecord(string name)
        {
            try
            {
                App?.UndoRecord.StartCustomRecord(name);
            }
            catch
            {
                // UndoRecord not supported in older Word versions — silently ignore
            }
        }

        /// <summary>
        /// Ends the current custom undo record.
        /// </summary>
        /// <summary>
        /// Replace all occurrences of a word in the active document (whole-word, case-insensitive).
        /// </summary>
        public static void ReplaceAllInDocument(string findWord, string replaceWord)
        {
            if (!IsDocumentOpen()) return;
            var doc = ActiveDoc;
            if (doc?.Content == null) return;

            BeginUndoRecord("Барчасини алмаштириш");
            try
            {
                var find = doc.Content.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = findWord;
                find.Replacement.Text = replaceWord;
                find.Forward = true;
                find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                find.MatchWholeWord = true;
                find.MatchCase = false;
                find.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
            }
            finally
            {
                EndUndoRecord();
            }
        }

        public static void EndUndoRecord()
        {
            try
            {
                App?.UndoRecord.EndCustomRecord();
            }
            catch { }
        }

        /// <summary>
        /// Checks if a document is currently open and active.
        /// </summary>
        public static bool IsDocumentOpen()
        {
            try { return App != null && App.Documents.Count > 0 && ActiveDoc != null; }
            catch { return false; }
        }

        /// <summary>
        /// Gets the filename of the active document (without path).
        /// </summary>
        public static string GetActiveDocumentName()
        {
            try { return ActiveDoc?.Name ?? "Номсиз ҳужжат"; }
            catch { return "Номсиз ҳужжат"; }
        }

        /// <summary>
        /// Creates a new blank Word document and returns it.
        /// </summary>
        public static Word.Document CreateNewDocument()
        {
            return App?.Documents.Add();
        }
    }
}
