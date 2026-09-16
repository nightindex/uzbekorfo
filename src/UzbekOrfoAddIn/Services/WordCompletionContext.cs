using System;
using System.Drawing;
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
                    document.TrackRevisions || document.ContentControls.Count != 0 ||
                    (bool)selection.get_Information(Word.WdInformation.wdWithInTable) ||
                    (bool)selection.get_Information(Word.WdInformation.wdInFieldCode) ||
                    (bool)selection.get_Information(Word.WdInformation.wdInFieldResult)) return null;
                int end = selection.End;
                if (end < 2) return null;
                range = selection.Range.Duplicate;
                range.SetRange(Math.Max(0, end - 81), end);
                string preceding = range.Text ?? "";
                int start = preceding.Length;
                while (start > 0 && WordCompletionEngine.IsTokenCharacter(preceding[start - 1])) start--;
                string prefix = preceding.Substring(start);
                if (prefix.Length < 2 || !WordCompletionEngine.IsToken(prefix)) return null;
                range.SetRange(end, end + 1);
                var following = range.Text;
                if (!string.IsNullOrEmpty(following) && WordCompletionEngine.IsTokenCharacter(following[0])) return null;
                range.SetRange(end - prefix.Length, end);
                if (range.Fields.Count != 0) return null;
                var result = new WordCompletionContext {
                    _range = range, Prefix = prefix, Start = range.Start, End = end, WindowHandle = app.ActiveWindow.Hwnd
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
                    _range.Start == Start && _range.End == End && _range.Text == Prefix &&
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

        public bool TryInsert(Word.Application app, string completion, bool requireFocus = true)
        {
            if (completion == null || !completion.StartsWith(Prefix, StringComparison.Ordinal) || completion.Length <= Prefix.Length ||
                !WordCompletionEngine.IsToken(completion)) return false;
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

        public Rectangle GetCaretBounds(Word.Application app)
        {
            return GetCaretAnchor(app).CaretBounds;
        }

        internal OverlayAnchor GetCaretAnchor(Word.Application app)
        {
            Word.Range caret = _range.Duplicate;
            Word.Font wordFont = null;
            Word.Window window = null;
            Word.View view = null;
            Word.Zoom zoom = null;
            try
            {
                caret.SetRange(End, End);
                int x, y, width, height;
                string fontName = "Segoe UI";
                float fontSize = 11f;
                int zoomPercentage = 100;
                FontStyle fontStyle = FontStyle.Regular;
                IntPtr owner = new IntPtr(WindowHandle);
                using (new DpiLayout.Context(DpiLayout.WindowContext(owner)))
                {
                    window = app.ActiveWindow;
                    window.GetPoint(out x, out y, out width, out height, caret);
                    Rectangle caretBounds;
                    bool nativeCaret = OverlayPositioner.TryGetNativeCaret(owner, out caretBounds);
                    if (!nativeCaret)
                    {
                        // A collapsed range can describe the paragraph end, whose
                        // height/format differs from the text the user just typed.
                        caret.SetRange(End - 1, End);
                        int textX, textY, textWidth, textHeight;
                        window.GetPoint(out textX, out textY, out textWidth, out textHeight, caret);
                        caretBounds = new Rectangle(x, textY, 1, Math.Max(1, textHeight));
                    }
                    // Read formatting from the last typed character, not the paragraph mark.
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
                    try
                    {
                        view = window.View;
                        zoom = view.Zoom;
                        int percentage = zoom.Percentage;
                        if (percentage >= 10 && percentage <= 500)
                            zoomPercentage = percentage;
                    }
                    catch (COMException) { /* Use 100% when Word cannot report zoom. */ }
                    // Word reports the document font size, not its zoomed screen size.
                    // Scale it so the ghost continuation visually joins the typed text.
                    fontSize *= zoomPercentage / 100f;
                    // Keep all screen-coordinate APIs in Word's DPI context. Otherwise
                    // monitor work areas can be virtualized differently from GetPoint.
                    var anchor = OverlayPositioner.CreateAnchor(
                        caretBounds,
                        owner, fontName, fontSize, fontStyle);
                    return nativeCaret ? anchor : OverlayPositioner.WithTextHeight(anchor);
                }
            }
            finally
            {
                if (zoom != null) Marshal.ReleaseComObject(zoom);
                if (view != null) Marshal.ReleaseComObject(view);
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
