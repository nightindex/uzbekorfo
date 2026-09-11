using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Services;
using Word = Microsoft.Office.Interop.Word;

internal static class Program
{
    [STAThread]
    private static int Main()
    {
        Word.Application app = null;
        var documents = new List<Word.Document>();
        bool ownsApplication = false;
        try
        {
            app = new Word.Application();
            // Never close or alter an application that already has a user's documents.
            if (app.Documents.Count != 0) throw new InvalidOperationException("Expected a new empty Word instance.");
            ownsApplication = true;
            app.Visible = false;
            app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            var a = app.Documents.Add(); documents.Add(a);
            var b = app.Documents.Add(); documents.Add(b);
            a.Content.Text = "x kiotb end";
            b.Content.Text = "x other end";
            var error = MakeError(a.Range(2, 7));
            var workflow = new ReplaceAllWorkflowService(null);
            Require(workflow.ApplyAll(b, new[] { error }) == 0, "Other document rejected");
            Require(b.Content.Text.TrimEnd('\r') == "x other end", "Other document unchanged");

            a.Range(0, 0).InsertBefore("new ");
            Require(error.Range.Text == "kiotb" && error.Range.Start == 6, "Word moved the live range");
            Require(workflow.ApplyAll(a, new[] { error }) == 1, "Correction follows live range");
            Require(a.Content.Text.TrimEnd('\r') == "new x kitob end", "Correct occurrence replaced");

            var stale = MakeError(a.Range(6, 11));
            a.Range(6, 11).Text = "other";
            Require(!DocumentHelper.TryReplaceError(stale, a, "kitob"), "Edited error rejected");

            a.Content.Text = "user error";
            var existing = a.Range(0, 4);
            existing.Underline = Word.WdUnderline.wdUnderlineWavy;
            existing.Font.UnderlineColor = Word.WdColor.wdColorRed;
            var marked = a.Range(5, 10);
            marked.Underline = Word.WdUnderline.wdUnderlineNone;
            marked.Font.UnderlineColor = Word.WdColor.wdColorAutomatic;
            var highlights = typeof(ErrorEntry).Assembly.GetType("UzbekOrfoAddIn.Services.DocumentHighlightService", true);
            highlights.GetMethod("Highlight").Invoke(null, new object[] { existing, Word.WdColor.wdColorGreen });
            highlights.GetMethod("Highlight").Invoke(null, new object[] { marked, Word.WdColor.wdColorRed });
            Require(marked.Underline == Word.WdUnderline.wdUnderlineWavy, "Owned mark applied");
            a.Range(5, 6).Underline = Word.WdUnderline.wdUnderlineSingle;
            highlights.GetMethod("ClearDocument").Invoke(null, new object[] { a });
            Require(existing.Underline == Word.WdUnderline.wdUnderlineWavy &&
                existing.Font.UnderlineColor == Word.WdColor.wdColorRed, "Existing underline preserved");
            Require(a.Range(5, 6).Underline == Word.WdUnderline.wdUnderlineSingle, "Later user formatting preserved");
            Require(a.Range(6, 10).Underline == Word.WdUnderline.wdUnderlineNone, "Owned mark cleared");

            a.Content.Text = "salom ";
            a.Activate();
            a.Range(0, 1).Select();
            Require(DocumentHelper.GetTargetRange(app.Selection, a).Text == "s", "Single-character selection remains selected");
            a.Range(5, 6).Select();
            Require(DocumentHelper.GetTargetRange(app.Selection, a).Text == " ", "Whitespace selection remains selected");
            DocumentHelper.ReplaceRangeText(a.Content, "салом \r", false);
            Require(a.Content.Text == "салом \r", "Exact transformation preserves trailing whitespace once: " + a.Content.Text.Replace("\r", "\\r"));
            var closed = MakeError(b.Range(2, 7));
            b.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            documents.Remove(b);
            Require(!DocumentHelper.TryReplaceError(closed, a, "kitob"), "Closed document range rejected");
            Console.WriteLine("PASS: real Word document ownership, live ranges, stale corrections, underline restoration, and exact replacement");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            foreach (var document in documents)
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            if (ownsApplication)
                try { app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            if (app != null) try { Marshal.ReleaseComObject(app); } catch { }
        }
    }

    private static ErrorEntry MakeError(Word.Range range) => new ErrorEntry
    {
        Range = range, Word = range.Text, OriginalText = range.Text,
        StartIndex = range.Start, EndIndex = range.End,
        Suggestions = new List<Suggestion> { new Suggestion("kitob", 1, 1) }
    };
    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
