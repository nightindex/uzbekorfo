using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.UI;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>UI-thread-only token snapshot. Owns only its duplicated Word range.</summary>
    public sealed class WordCompletionContext : IDisposable
    {
        private Word.Range _range;
        public string Prefix { get; private set; }
        public string PrecedingContext { get; private set; }
        public int Start { get; private set; }
        public int End { get; private set; }
        public int WindowHandle { get; private set; }

        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetFocus();
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int count);
        [DllImport("imm32.dll")] private static extern IntPtr ImmGetContext(IntPtr hwnd);
        [DllImport("imm32.dll")] private static extern bool ImmReleaseContext(IntPtr hwnd, IntPtr context);
        [DllImport("imm32.dll", CharSet = CharSet.Unicode)]
        private static extern int ImmGetCompositionStringW(IntPtr context, int index, IntPtr buffer, int length);

        public static bool HasEditingFocus(Word.Application app)
        {
            try
            {
                return HasEditingFocus(app.ActiveWindow.Hwnd);
            }
            catch (COMException) { return false; }
        }

        /// <summary>Native-only focus guard safe to use from the existing Word keyboard hook.</summary>
        public static bool HasEditingFocus(int windowHandle)
        {
                if (GetForegroundWindow() != new IntPtr(windowHandle)) return false;
                var focus = GetFocus();
                var name = new StringBuilder(64);
                GetClassName(focus, name, name.Capacity);
                if (name.ToString() != "_WwG") return false;
                var context = ImmGetContext(focus);
                try { return context == IntPtr.Zero || ImmGetCompositionStringW(context, 8, IntPtr.Zero, 0) <= 0; }
                finally { if (context != IntPtr.Zero) ImmReleaseContext(focus, context); }
        }

        public static WordCompletionContext Capture(Word.Application app, bool requireFocus = true)
        {
            Word.Range range = null;
            try
            {
                if (requireFocus && !HasEditingFocus(app)) return null;
                var selection = app.Selection;
                var document = app.ActiveDocument;
                if (selection.Start != selection.End || selection.StoryType != Word.WdStoryType.wdMainTextStory ||
                    document.ReadOnly || document.ProtectionType != Word.WdProtectionType.wdNoProtection ||
                    document.TrackRevisions ||
                    (bool)selection.get_Information(Word.WdInformation.wdWithInTable) ||
                    (bool)selection.get_Information(Word.WdInformation.wdInFieldCode) ||
                    (bool)selection.get_Information(Word.WdInformation.wdInFieldResult)) return null;
                Word.ContentControl parentControl = selection.ParentContentControl;
                if (parentControl != null)
                {
                    Marshal.ReleaseComObject(parentControl);
                    return null;
                }
                int end = selection.End;
                if (end < 2) return null;
                range = selection.Range.Duplicate;
                range.SetRange(Math.Max(0, end - 512), end);
                string preceding = range.Text ?? "";
                int start = preceding.Length;
                while (start > 0 && WordCompletionEngine.IsTokenCharacter(preceding[start - 1])) start--;
                string prefix = preceding.Substring(start);
                if (prefix.Length > 0 && !WordCompletionEngine.IsToken(prefix)) return null;
                if (prefix.Length == 0 && (preceding.Length == 0 || preceding[preceding.Length - 1] != ' ')) return null;
                string context = preceding.Substring(0, start);
                range.SetRange(end, end + 1);
                var following = range.Text;
                if (!string.IsNullOrEmpty(following) && WordCompletionEngine.IsTokenCharacter(following[0])) return null;
                range.SetRange(end - prefix.Length, end);
                Word.Fields fields = null;
                Word.ContentControls controls = null;
                try
                {
                    fields = range.Fields;
                    controls = range.ContentControls;
                    if (fields.Count != 0 || controls.Count != 0) return null;
                }
                finally
                {
                    if (controls != null) Marshal.ReleaseComObject(controls);
                    if (fields != null) Marshal.ReleaseComObject(fields);
                }
                var result = new WordCompletionContext {
                    _range = range, Prefix = prefix, PrecedingContext = context,
                    Start = range.Start, End = end, WindowHandle = app.ActiveWindow.Hwnd
                };
                range = null;
                return result;
            }
            catch (COMException) { return null; }
            finally { if (range != null) Marshal.ReleaseComObject(range); }
        }

        public bool SameAs(WordCompletionContext other)
        {
            try
            {
                return other != null && _range != null && other._range != null &&
                    WindowHandle == other.WindowHandle && Start == other.Start && End == other.End && Prefix == other.Prefix &&
                    PrecedingContext == other.PrecedingContext &&
                    _range.Start == Start && _range.End == End && (_range.Text ?? string.Empty) == Prefix &&
                    DocumentHelper.IsSameDocument(_range.Document, other._range.Document);
            }
            catch (COMException) { return false; }
        }

        internal bool CanContinueAt(WordCompletionContext other)
        {
            try
            {
                return other != null && _range != null && other._range != null &&
                    WindowHandle == other.WindowHandle && Start == other.Start && End < other.End &&
                    other.Prefix.StartsWith(Prefix, StringComparison.Ordinal) &&
                    DocumentHelper.IsSameDocument(_range.Document, other._range.Document);
            }
            catch (COMException) { return false; }
        }

        /// <summary>Advance an already visible phrase only through exact matching typing.</summary>
        public bool TryContinueSuggestion(WordCompletionContext other, string completion, out string continued)
        {
            continued = null;
            try
            {
                if (other == null || _range == null || other._range == null || WindowHandle != other.WindowHandle ||
                    !DocumentHelper.IsSameDocument(_range.Document, other._range.Document) ||
                    completion == null || !completion.StartsWith(Prefix, StringComparison.Ordinal)) return false;
                int typed = other.End - End;
                string tail = completion.Substring(Prefix.Length);
                if (typed <= 0 || typed >= tail.Length) return false;
                string expected = PrecedingContext + Prefix + tail.Substring(0, typed);
                if (expected.Length > 512) expected = expected.Substring(expected.Length - 512);
                if (other.PrecedingContext + other.Prefix != expected) return false;
                continued = other.Prefix + tail.Substring(typed);
                return IsSafeCompletion(continued);
            }
            catch (COMException) { return false; }
        }

        public bool TryInsert(Word.Application app, string completion, bool requireFocus = true)
        {
            if (completion == null || !completion.StartsWith(Prefix, StringComparison.Ordinal) || completion.Length <= Prefix.Length ||
                !IsSafeCompletion(completion)) return false;
            using (var current = Capture(app, requireFocus))
            {
                if (!SameAs(current)) return false;
                // One explicit edit, preserving typed apostrophes/case and adjacent formatting.
                var undo = app.UndoRecord;
                undo.StartCustomRecord("MatnAI таклифи");
                try
                {
                    _range.SetRange(End, End);
                    _range.Text = completion.Substring(Prefix.Length);
                    _range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    _range.Select();
                }
                finally { undo.EndCustomRecord(); }
                return true;
            }
        }

        internal static bool IsSafeCompletion(string completion)
        {
            if (string.IsNullOrEmpty(completion) || completion.Length > 404) return false;
            var words = completion.Split(' ');
            return words.Length <= 5 && words.All(WordCompletionEngine.IsToken);
        }

        public Rectangle GetCaretBounds(Word.Application app)
        {
            return GetCaretAnchor(app).CaretBounds;
        }

        private static bool TryGetTypingFont(Word.Application app, int end,
            out string name, out float size, out FontStyle style)
        {
            name = null;
            size = 0;
            style = FontStyle.Regular;
            Word.Selection selection = null;
            Word.Font font = null;
            try
            {
                selection = app.Selection;
                if (selection.Start != end || selection.End != end) return false;
                font = selection.Font;
                string currentName = font.Name;
                float currentSize = font.Size;
                if (string.IsNullOrWhiteSpace(currentName) || currentSize < 6f || currentSize > 96f)
                    return false;
                name = currentName;
                size = currentSize;
                if (font.Bold == -1) style |= FontStyle.Bold;
                if (font.Italic == -1) style |= FontStyle.Italic;
                return true;
            }
            catch (COMException) { return false; }
            finally
            {
                if (font != null) Marshal.ReleaseComObject(font);
                if (selection != null) Marshal.ReleaseComObject(selection);
            }
        }

        private static int GetZoomPercentage(Word.Window window)
        {
            Word.View view = null;
            Word.Zoom zoom = null;
            try
            {
                view = window.View;
                zoom = view.Zoom;
                int percentage = zoom.Percentage;
                return percentage >= 10 && percentage <= 500 ? percentage : 100;
            }
            catch (COMException) { return 100; }
            finally
            {
                if (zoom != null) Marshal.ReleaseComObject(zoom);
                if (view != null) Marshal.ReleaseComObject(view);
            }
        }

        internal bool MatchesPresentationFormatting(Word.Application app, OverlayAnchor anchor)
        {
            if (anchor == null || _range == null) return false;
            string name;
            float size;
            FontStyle style;
            if (!TryGetTypingFont(app, End, out name, out size, out style)) return false;
            Word.Window window = null;
            try
            {
                window = app.ActiveWindow;
                float effectiveSize = size * GetZoomPercentage(window) / 100f;
                return string.Equals(name, anchor.FontName, StringComparison.OrdinalIgnoreCase) &&
                    Math.Abs(effectiveSize - anchor.FontSizePoints) < 0.01f &&
                    style == anchor.FontStyle;
            }
            catch (COMException) { return false; }
            finally { if (window != null) Marshal.ReleaseComObject(window); }
        }

        internal OverlayAnchor GetCaretAnchor(Word.Application app)
        {
            Word.Range caret = _range.Duplicate;
            Word.Font wordFont = null;
            Word.Window window = null;
            try
            {
                caret.SetRange(End, End);
                string fontName = "Segoe UI";
                float fontSize = 11f;
                int zoomPercentage = 100;
                FontStyle fontStyle = FontStyle.Regular;
                IntPtr owner = new IntPtr(WindowHandle);
                using (new DpiLayout.Context(DpiLayout.WindowContext(owner)))
                {
                    window = app.ActiveWindow;
                    Rectangle caretBounds;
                    bool nativeCaret = OverlayPositioner.TryGetNativeCaret(owner, out caretBounds);
                    if (!nativeCaret)
                    {
                        int x, y, width, height;
                        window.GetPoint(out x, out y, out width, out height, caret);
                        // A collapsed range can describe the paragraph end, whose
                        // height/format differs from the text the user just typed.
                        caret.SetRange(End - 1, End);
                        int textX, textY, textWidth, textHeight;
                        window.GetPoint(out textX, out textY, out textWidth, out textHeight, caret);
                        caretBounds = new Rectangle(x, textY, 1, Math.Max(1, textHeight));
                    }
                    // The collapsed selection carries the formatting of the next
                    // character. It may differ from the preceding character after
                    // a font-size or style change at the caret.
                    string typingFontName;
                    float typingFontSize;
                    FontStyle typingFontStyle;
                    if (TryGetTypingFont(app, End, out typingFontName, out typingFontSize, out typingFontStyle))
                    {
                        fontName = typingFontName;
                        fontSize = typingFontSize;
                        fontStyle = typingFontStyle;
                    }
                    else
                    {
                        // Retain the last character as a fallback when Word cannot
                        // report insertion formatting for this selection.
                        caret.SetRange(End - 1, End);
                        try
                        {
                            wordFont = caret.Font;
                            if (!string.IsNullOrWhiteSpace(wordFont.Name)) fontName = wordFont.Name;
                            if (wordFont.Size >= 6f && wordFont.Size <= 96f) fontSize = wordFont.Size;
                            if (wordFont.Bold == -1) fontStyle |= FontStyle.Bold;
                            if (wordFont.Italic == -1) fontStyle |= FontStyle.Italic;
                        }
                        catch (COMException) { /* Use safe font defaults. */ }
                    }
                    zoomPercentage = GetZoomPercentage(window);
                    // Word reports the document font size, not its zoomed screen size.
                    // Scale it so the ghost continuation visually joins the typed text.
                    fontSize *= zoomPercentage / 100f;
                    // Keep all screen-coordinate APIs in Word's DPI context. Otherwise
                    // monitor work areas can be virtualized differently from GetPoint.
                    var anchor = OverlayPositioner.CreateAnchor(
                        caretBounds,
                        owner, fontName, fontSize, fontStyle);
                    Word.PageSetup page = null;
                    Word.ParagraphFormat paragraph = null;
                    Word.TextColumns columns = null;
                    try
                    {
                        page = _range.Document.PageSetup;
                        paragraph = _range.ParagraphFormat;
                        columns = page.TextColumns;
                        caret.SetRange(End, End);
                        float horizontal = Convert.ToSingle(caret.get_Information(Word.WdInformation.wdHorizontalPositionRelativeToPage));
                        float remaining = page.PageWidth - page.RightMargin - Math.Max(0, paragraph.RightIndent) - horizontal;
                        // Inline placement is conservative when Word cannot expose a single text column.
                        anchor.TextRight = horizontal < 0 || columns.Count != 1 ? caretBounds.Right :
                            caretBounds.Right + Math.Max(0, (int)(remaining * zoomPercentage / 100f * anchor.Dpi / 72f));
                    }
                    catch (COMException) { anchor.TextRight = caretBounds.Right; }
                    finally
                    {
                        if (columns != null) Marshal.ReleaseComObject(columns);
                        if (paragraph != null) Marshal.ReleaseComObject(paragraph);
                        if (page != null) Marshal.ReleaseComObject(page);
                    }
                    return nativeCaret ? anchor : OverlayPositioner.WithTextHeight(anchor);
                }
            }
            finally
            {
                if (wordFont != null) Marshal.ReleaseComObject(wordFont);
                if (window != null) Marshal.ReleaseComObject(window);
                Marshal.ReleaseComObject(caret);
            }
        }
        public void Dispose()
        {
            if (_range == null) return;
            Marshal.ReleaseComObject(_range);
            _range = null;
        }
    }
}
