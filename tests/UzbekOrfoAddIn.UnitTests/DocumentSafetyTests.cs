using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Services;
using Word = Microsoft.Office.Interop.Word;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class DocumentSafetyTests : IDisposable
{
    public void Dispose() => DocumentHighlightService.ClearAll();

    private static ErrorEntry Error(Word.Document doc, int start, int end) => new()
    {
        Word = doc.Range(start, end).Text, OriginalText = doc.Range(start, end).Text,
        Range = doc.Range(start, end), StartIndex = start, EndIndex = end,
        Suggestions = new List<Suggestion> { new("kitob", 1, 1) }
    };

    [Fact]
    public void ReplaceAll_RejectsAnotherDocumentEvenWithSameNameAndText()
    {
        var first = new Word.Document("kiotb");
        var second = new Word.Document("kiotb");
        var error = Error(first, 0, 5);
        Assert.Equal(0, new ReplaceAllWorkflowService(null).ApplyAll(second, new[] { error }));
        Assert.Equal("kiotb", second.Content.Text);
        Assert.False(error.IsResolved);
    }

    [Fact]
    public void ReplaceAll_UsesLiveRangeAfterPrecedingInsertion()
    {
        var doc = new Word.Document("kiotb");
        var error = Error(doc, 0, 5);
        // Model Word moving the live range after text is inserted before the error.
        doc.Content.Text = "new kiotb";
        error.Range.Start = 4;
        error.Range.End = 9;
        Assert.Equal(1, new ReplaceAllWorkflowService(null).ApplyAll(doc, new[] { error }));
        Assert.Equal("new kitob", doc.Content.Text);
    }

    [Fact]
    public void ReplaceAll_SkipsEditedTextAndStillFixesValidErrors()
    {
        var doc = new Word.Document("kiotb kiotb");
        var changed = Error(doc, 0, 5);
        var valid = Error(doc, 6, 11);
        doc.Range(0, 5).Text = "other";
        Assert.Equal(1, new ReplaceAllWorkflowService(null).ApplyAll(doc, new[] { changed, valid }));
        Assert.Equal("other kitob", doc.Content.Text);
        Assert.False(changed.IsResolved);
    }

    [Theory]
    [InlineData("a", Word.WdSelectionType.wdSelectionNormal)]
    [InlineData(" ", Word.WdSelectionType.wdSelectionNormal)]
    [InlineData("a", Word.WdSelectionType.wdSelectionRow)]
    public void TargetRange_PreservesSingleCharacterAndWhitespaceSelections(string selected, Word.WdSelectionType type)
    {
        var doc = new Word.Document(selected + " rest");
        Globals.ThisAddIn.Application.ActiveDocument = doc;
        Globals.ThisAddIn.Application.Selection = new Word.Selection { Type = type, Range = doc.Range(0, 1) };
        Assert.Equal(selected, DocumentHelper.GetTargetRange().Text);
    }

    [Fact]
    public void ExactReplacement_DoesNotDuplicateTrailingWhitespace()
    {
        var doc = new Word.Document("salom \r");
        DocumentHelper.ReplaceRangeText(doc.Content, "салом \r", preserveTrailingWhitespace: false);
        Assert.Equal("салом \r", doc.Content.Text);
    }

    [Fact]
    public void Highlights_PreserveExistingUnderlinesAndOnlyClearOwnedMarks()
    {
        var doc = new Word.Document("user error");
        var user = doc.Range(0, 4);
        user.Underline = Word.WdUnderline.wdUnderlineWavy;
        user.Font.UnderlineColor = Word.WdColor.wdColorRed;
        DocumentHighlightService.Highlight(user, Word.WdColor.wdColorGreen);
        var error = doc.Range(5, 10);
        DocumentHighlightService.Highlight(error, Word.WdColor.wdColorRed);
        Assert.Equal(Word.WdUnderline.wdUnderlineWavy, error.Underline);
        DocumentHighlightService.ClearDocument(doc);
        Assert.Equal(Word.WdUnderline.wdUnderlineWavy, user.Underline);
        Assert.Equal(Word.WdColor.wdColorRed, user.Font.UnderlineColor);
        Assert.Equal(Word.WdUnderline.wdUnderlineNone, error.Underline);
        Assert.Equal(Word.WdColor.wdColorAutomatic, error.Font.UnderlineColor);
    }

    [Fact]
    public void HighlightCleanup_PreservesUserFormattingAppliedAfterCheck()
    {
        var doc = new Word.Document("error");
        DocumentHighlightService.Highlight(doc.Content, Word.WdColor.wdColorRed);
        doc.Range(0, 1).Underline = Word.WdUnderline.wdUnderlineSingle;
        DocumentHighlightService.ClearDocument(doc);
        Assert.Equal(Word.WdUnderline.wdUnderlineSingle, doc.Range(0, 1).Underline);
        Assert.Equal(Word.WdUnderline.wdUnderlineNone, doc.Range(1, 5).Underline);
    }

    [Fact]
    public void ReplacingMarkedWord_DoesNotTransferHighlightToReplacement()
    {
        var doc = new Word.Document("kiotb");
        var error = Error(doc, 0, 5);
        DocumentHighlightService.Highlight(error.Range, Word.WdColor.wdColorRed);
        Assert.True(DocumentHelper.TryReplaceError(error, doc, "kitob"));
        Assert.Equal(Word.WdUnderline.wdUnderlineNone, doc.Content.Underline);
    }
}
