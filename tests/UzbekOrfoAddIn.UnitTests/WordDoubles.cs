// Small in-memory host double. Exercises production workflow code without Word.
// Does not claim to reproduce COM, Word's layout, undo, or all live-range semantics.
#nullable disable
using System.Collections.Generic;
using System.Linq;

namespace UzbekOrfoAddIn
{
    internal static class Globals
    {
        public static TestAddIn ThisAddIn { get; } = new TestAddIn();
    }
    internal sealed class TestAddIn
    {
        public Microsoft.Office.Interop.Word.Application Application { get; set; } = new Microsoft.Office.Interop.Word.Application();
    }
}
namespace UzbekOrfoAddIn.Helpers
{
    internal static class Logger
    {
        public static void Info(string text) { }
        public static void Warn(string text) { }
        public static void Error(string text, System.Exception exception) { }
    }
}
namespace Microsoft.Office.Interop.Word
{
    public enum WdSelectionType { wdNoSelection, wdSelectionIP, wdSelectionNormal, wdSelectionRow }
    public enum WdUnderline { wdUnderlineNone, wdUnderlineSingle, wdUnderlineWavy }
    public enum WdColor { wdColorAutomatic = -1, wdColorRed = 255, wdColorGreen = 32768 }
    public enum WdConstants { wdUndefined = 9999999 }
    public enum WdFindWrap { wdFindContinue }
    public enum WdStoryType { wdMainTextStory }
    public enum WdReplace { wdReplaceAll }
    public sealed class Application
    {
        public Document ActiveDocument { get; set; }
        public Selection Selection { get; set; }
        public Documents Documents { get; } = new Documents();
        public UndoRecord UndoRecord { get; } = new UndoRecord();
        public Window ActiveWindow => throw new System.NotSupportedException();
    }
    public sealed class Documents
    {
        public int Count => 1;
        public Document Add() => throw new System.NotSupportedException();
    }
    public sealed class UndoRecord
    {
        public void StartCustomRecord(string name) { }
        public void EndCustomRecord() { }
    }
    public sealed class Window { public void ScrollIntoView(Range range) => throw new System.NotSupportedException(); }
    public sealed class Selection
    {
        public WdSelectionType Type { get; set; }
        public Range Range { get; set; }
        public string Text => Range.Text;
    }
    public sealed class Document
    {
        internal sealed class Character
        {
            public char Text;
            public WdUnderline Underline;
            public WdColor Color = WdColor.wdColorAutomatic;
        }
        internal List<Character> Cells;
        internal List<Range> Ranges = new List<Range>();
        public Document(string text) { Cells = text.Select(c => new Character { Text = c }).ToList(); }
        public string Name { get; set; } = "Document1";
        public Range Content => Range(0, Cells.Count);
        public Range Range(int start, int end)
        {
            var range = new Range(this, start, end);
            Ranges.Add(range);
            return range;
        }
        internal void Replace(int start, int end, string text)
        {
            int delta = text.Length - (end - start);
            var formatting = Cells.Count > start ? Cells[start] : new Character();
            Cells.RemoveRange(start, end - start);
            Cells.InsertRange(start, text.Select(c => new Character { Text = c, Underline = formatting.Underline, Color = formatting.Color }));
            foreach (var range in Ranges)
            {
                if (range.Start >= end && range.Start != start) range.Start += delta;
                if (range.End >= end) range.End += delta;
            }
        }
    }
    public sealed class Range
    {
        internal Range(Document document, int start, int end) { Document = document; Start = start; End = end; }
        public Document Document { get; }
        public int Start { get; set; }
        public int End { get; set; }
        internal IEnumerable<Document.Character> Cells => Document.Cells.Skip(Start).Take(End - Start);
        public string Text
        {
            get => new string(Cells.Select(c => c.Text).ToArray());
            set => Document.Replace(Start, End, value);
        }
        public Range Duplicate => Document.Range(Start, End);
        public WdStoryType StoryType => WdStoryType.wdMainTextStory;
        public WdUnderline Underline
        {
            get { var styles = Cells.Select(c => c.Underline).Distinct().ToArray(); return styles.Length == 1 ? styles[0] : (WdUnderline)WdConstants.wdUndefined; }
            set { foreach (var c in Cells) c.Underline = value; }
        }
        public Font Font => new Font(this);
        public Characters Characters => new Characters(this);
        public bool InStory(Range other) => ReferenceEquals(Document, other.Document);
        public Find Find => throw new System.NotSupportedException();
        public Dictionary<int, Paragraph> Paragraphs => throw new System.NotSupportedException();
        public Dictionary<int, Range> Words => throw new System.NotSupportedException();
        public void Select() => throw new System.NotSupportedException();
    }
    public sealed class Characters
    {
        private readonly Range _range;
        public Characters(Range range) { _range = range; }
        public int Count => _range.End - _range.Start;
        public Range this[int index] => _range.Document.Range(_range.Start + index - 1, _range.Start + index);
    }
    public sealed class Font
    {
        private readonly Range _range;
        public Font(Range range) { _range = range; }
        public WdColor UnderlineColor
        {
            get { var colors = _range.Cells.Select(c => c.Color).Distinct().ToArray(); return colors.Length == 1 ? colors[0] : (WdColor)WdConstants.wdUndefined; }
            set { foreach (var c in _range.Cells) c.Color = value; }
        }
    }
    public sealed class Paragraph { public Range Range => throw new System.NotSupportedException(); }
    public sealed class Find
    {
        public string Text { get; set; }
        public Find Replacement => this;
        public bool Forward { get; set; }
        public WdFindWrap Wrap { get; set; }
        public bool MatchWholeWord { get; set; }
        public bool MatchCase { get; set; }
        public void ClearFormatting() => throw new System.NotSupportedException();
        public void Execute(WdReplace Replace) => throw new System.NotSupportedException();
    }
}
