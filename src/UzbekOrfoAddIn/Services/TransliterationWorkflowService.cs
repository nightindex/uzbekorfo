using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.UI;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates transliteration commands for ribbon actions.
    /// </summary>
    public sealed class TransliterationWorkflowService
    {
        private readonly ITransliterator _transliterator;
        private readonly UzbekApostropheService _apostropheService;

        public TransliterationWorkflowService(
            ITransliterator transliterator,
            UzbekApostropheService apostropheService)
        {
            _transliterator = transliterator;
            _apostropheService = apostropheService ?? new UzbekApostropheService();
        }

        public TransliterationWorkflowResult ExecuteLatinToCyrillic()
        {
            if (!DocumentHelper.IsDocumentOpen())
                return TransliterationWorkflowResult.NoOp();

            if (_transliterator == null)
                return TransliterationWorkflowResult.ServiceUnavailable();

            var range = DocumentHelper.GetTargetRange();
            if (range == null)
                return TransliterationWorkflowResult.NoOp();

            string sourceText = range.Text;
            if (string.IsNullOrWhiteSpace(sourceText))
                return TransliterationWorkflowResult.TextNotFound();

            var script = _transliterator.DetectScript(sourceText);
            if (script == ScriptType.Cyrillic)
            {
                if (!ModernMessageBox.Confirm("Матн аллақачон кириллда. Бари бир ўтказайми?",
                        "Лотиндан Кириллга"))
                {
                    return TransliterationWorkflowResult.Cancelled();
                }
            }
            else
            {
                if (!ModernMessageBox.Confirm($"{DescribeScope(range)} Лотиндан Кириллга алмаштирасизми?",
                        "Лотиндан Кириллга"))
                {
                    return TransliterationWorkflowResult.Cancelled();
                }
            }

            int wordCount = 0;
            RunWithBusyIndicator(() =>
            {
                DocumentHelper.BeginUndoRecord("Лотиндан Кириллга");
                try
                {
                    wordCount = ReplaceTextByToken(range, _transliterator.ToCyrillic);
                }
                finally { DocumentHelper.EndUndoRecord(); }
            });

            return TransliterationWorkflowResult.Completed("Кирилл", wordCount);
        }

        public TransliterationWorkflowResult ExecuteCyrillicToLatin()
        {
            if (!DocumentHelper.IsDocumentOpen())
                return TransliterationWorkflowResult.NoOp();

            if (_transliterator == null)
                return TransliterationWorkflowResult.ServiceUnavailable();

            var range = DocumentHelper.GetTargetRange();
            if (range == null)
                return TransliterationWorkflowResult.NoOp();

            string sourceText = range.Text;
            if (string.IsNullOrWhiteSpace(sourceText))
                return TransliterationWorkflowResult.TextNotFound();

            var script = _transliterator.DetectScript(sourceText);
            if (script == ScriptType.Latin)
            {
                if (!ModernMessageBox.Confirm("Матн аллақачон лотинда. Бари бир ўтказайми?",
                        "Кириллдан Лотинга"))
                {
                    return TransliterationWorkflowResult.Cancelled();
                }
            }
            else
            {
                if (!ModernMessageBox.Confirm($"{DescribeScope(range)} Кириллдан Лотинга алмаштирасизми?",
                        "Кириллдан Лотинга"))
                {
                    return TransliterationWorkflowResult.Cancelled();
                }
            }

            int wordCount = 0;
            RunWithBusyIndicator(() =>
            {
                DocumentHelper.BeginUndoRecord("Кириллдан Лотинга");
                try
                {
                    wordCount = ReplaceTextByToken(range,
                        text => _apostropheService.NormalizeText(_transliterator.ToLatin(text)));
                }
                finally { DocumentHelper.EndUndoRecord(); }
            });

            return TransliterationWorkflowResult.Completed("Лотин", wordCount);
        }

        public TransliterationWorkflowResult ExecuteSwitchScript()
        {
            if (!DocumentHelper.IsDocumentOpen())
                return TransliterationWorkflowResult.NoOp();

            if (_transliterator == null)
                return TransliterationWorkflowResult.ServiceUnavailable();

            var range = DocumentHelper.GetTargetRange();
            if (range == null)
                return TransliterationWorkflowResult.NoOp();

            string sourceText = range.Text;
            if (string.IsNullOrWhiteSpace(sourceText))
                return TransliterationWorkflowResult.TextNotFound();

            var script = _transliterator.DetectScript(sourceText);
            string direction;
            Func<string, string> convert;

            switch (script)
            {
                case ScriptType.Cyrillic:
                    convert = text => _apostropheService.NormalizeText(_transliterator.ToLatin(text));
                    direction = "Кирилл → Лотин";
                    break;
                case ScriptType.Latin:
                    convert = _transliterator.ToCyrillic;
                    direction = "Лотин → Кирилл";
                    break;
                case ScriptType.Mixed:
                    var result = ModernMessageBox.ConfirmOrCancel(
                        "Матнда аралаш ёзув аниқланди.\nКириллга ўтказайми?",
                        "Ёзувни алмаштириш",
                        "Кириллга", "Лотинга", "Бекор");

                    if (result == DialogResult.Yes)
                    {
                        convert = _transliterator.ToCyrillic;
                        direction = "Аралаш → Кирилл";
                    }
                    else if (result == DialogResult.No)
                    {
                        convert = text => _apostropheService.NormalizeText(_transliterator.ToLatin(text));
                        direction = "Аралаш → Лотин";
                    }
                    else
                    {
                        return TransliterationWorkflowResult.Cancelled();
                    }
                    break;
                default:
                    return TransliterationWorkflowResult.ScriptUndetermined();
            }

            int wordCount = 0;
            RunWithBusyIndicator(() =>
            {
                DocumentHelper.BeginUndoRecord("Ёзувни алмаштириш");
                try { wordCount = ReplaceTextByToken(range, convert); }
                finally { DocumentHelper.EndUndoRecord(); }
            });

            return TransliterationWorkflowResult.Completed(direction, wordCount);
        }

        private static string DescribeScope(Microsoft.Office.Interop.Word.Range range)
        {
            return range.IsEqual(range.Document.Content) ? "Бутун ҳужжатни" : "Танланган матнни";
        }

        /// <summary>
        /// Replace only letter spans. Word's Words ranges can contain paragraph
        /// and table-cell marks, so replacing those ranges destroys document layout.
        /// </summary>
        private static int ReplaceTextByToken(Word.Range range, Func<string, string> convert)
        {
            if (range == null || convert == null) return 0;

            var replacements = new List<WordReplacement>();
            var convertedWords = new Dictionary<string, string>(StringComparer.Ordinal);
            Func<string, string> convertCached = source =>
            {
                if (!convertedWords.TryGetValue(source, out string target))
                {
                    target = convert(source);
                    convertedWords.Add(source, target);
                }
                return target;
            };
            Word.Document document = range.Document;
            int nextStart = range.Start;
            int searchEnd = range.End;
            int scanned = 0;
            Word.Range searchRange = null;
            Word.Find find = null;
            bool retriedSearch = false;
            var scanTimer = Stopwatch.StartNew();
            try
            {
                searchRange = range.Duplicate;
                find = searchRange.Find;
                ConfigureTransliterationFind(find, forward: true);
                while (nextStart < searchEnd)
                {
                    if (!find.Execute()) break;

                    int matchStart = searchRange.Start;
                    int matchEnd = searchRange.End;
                    // Word can report a match outside a bounded Range even with
                    // wdFindStop. Never modify text outside the requested range.
                    if (matchStart < nextStart || matchEnd <= matchStart || matchEnd > searchEnd)
                    {
                        if (!retriedSearch)
                        {
                            retriedSearch = true;
                            if (Marshal.IsComObject(find)) Marshal.ReleaseComObject(find);
                            find = null;
                            if (Marshal.IsComObject(searchRange)) Marshal.ReleaseComObject(searchRange);
                            searchRange = range.Duplicate;
                            searchRange.SetRange(nextStart, searchEnd);
                            find = searchRange.Find;
                            ConfigureTransliterationFind(find, forward: true);
                            continue;
                        }
                        Logger.Info($"Transliteration search reached range boundary after {scanned} matches (next {nextStart}, end {searchEnd}, found {matchStart}-{matchEnd}).");
                        scanned += CollectRemainingMatchesBackward(range, nextStart, searchEnd,
                            convertCached, replacements);
                        break;
                    }
                    retriedSearch = false;
                    string matchedText = searchRange.Text ?? string.Empty;
                    AddTokenReplacement(replacements, matchStart, matchedText, convertCached);
                    scanned++;

                    nextStart = Math.Max(matchEnd, matchStart + 1);
                    if (nextStart >= searchEnd) break;
                    searchRange.SetRange(nextStart, searchEnd);
                }
            }
            finally
            {
                if (find != null && Marshal.IsComObject(find))
                    Marshal.ReleaseComObject(find);
                if (searchRange != null && Marshal.IsComObject(searchRange))
                    Marshal.ReleaseComObject(searchRange);
            }

            DocumentHighlightService.ClearRange(range);
            long scanMilliseconds = scanTimer.ElapsedMilliseconds;
            var replaceTimer = Stopwatch.StartNew();
            replacements.Sort((left, right) => left.Start.CompareTo(right.Start));
            // Later edits cannot shift the document positions of earlier text.
            for (int index = replacements.Count - 1; index >= 0; index--)
            {
                var replacement = replacements[index];
                Word.Range word = null;
                try
                {
                    word = document.Range(replacement.Start, replacement.End);
                    word.Text = replacement.Text;
                }
                finally
                {
                    if (word != null && Marshal.IsComObject(word)) Marshal.ReleaseComObject(word);
                }
            }
            if (scanMilliseconds + replaceTimer.ElapsedMilliseconds >= 1000)
                Logger.Info($"Transliteration: {scanned} matches, {replacements.Count} edits; scan {scanMilliseconds} ms, replace {replaceTimer.ElapsedMilliseconds} ms.");
            return replacements.Count;
        }

        // Word performs this wildcard search natively. It returns exact ranges and
        // avoids thousands of slow per-character COM calls.
        private const string TransliterationFindPattern =
            "[A-Za-z\u0410-\u044F\u0401\u0451\u040E\u045E\u049A\u049B\u0492\u0493\u04B2\u04B3'\u02BB\u02BC\u2018\u2019]@";

        private static void ConfigureTransliterationFind(Word.Find find, bool forward)
        {
            find.ClearFormatting();
            find.Text = TransliterationFindPattern;
            find.Forward = forward;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.Format = false;
            find.MatchWildcards = true;
        }

        private static int CollectRemainingMatchesBackward(Word.Range originalRange, int start, int end,
            Func<string, string> convert, List<WordReplacement> replacements)
        {
            Word.Range searchRange = null;
            Word.Find find = null;
            int scanned = 0;
            try
            {
                searchRange = originalRange.Duplicate;
                searchRange.SetRange(start, end);
                find = searchRange.Find;
                ConfigureTransliterationFind(find, forward: false);
                int nextEnd = end;
                while (nextEnd > start && find.Execute())
                {
                    int matchStart = searchRange.Start;
                    int matchEnd = searchRange.End;
                    if (matchStart < start || matchEnd > nextEnd || matchEnd <= matchStart)
                        break;

                    AddTokenReplacement(replacements, matchStart, searchRange.Text ?? string.Empty, convert);
                    scanned++;
                    nextEnd = matchStart;
                    if (nextEnd <= start) break;
                    searchRange.SetRange(start, nextEnd);
                }
            }
            finally
            {
                if (find != null && Marshal.IsComObject(find)) Marshal.ReleaseComObject(find);
                if (searchRange != null && Marshal.IsComObject(searchRange))
                    Marshal.ReleaseComObject(searchRange);
            }
            return scanned;
        }

        private static void AddTokenReplacement(List<WordReplacement> replacements,
            int matchStart, string matchedText, Func<string, string> convert)
        {
            int index = 0;
            while (index < matchedText.Length)
            {
                while (index < matchedText.Length && !IsWordLetter(matchedText[index])) index++;
                if (index >= matchedText.Length) return;

                int start = index++;
                while (index < matchedText.Length &&
                    (IsWordLetter(matchedText[index]) || IsInternalApostrophe(matchedText, index)))
                    index++;

                string source = matchedText.Substring(start, index - start);
                string target = convert(source);
                if (!string.Equals(source, target, StringComparison.Ordinal))
                    replacements.Add(new WordReplacement(matchStart + start, matchStart + index, target));
            }
        }

        private static bool IsWordLetter(char value)
        {
            return char.IsLetter(value) && !IsApostrophe(value);
        }

        private static bool IsInternalApostrophe(string text, int index)
        {
            return index > 0 && index + 1 < text.Length && IsApostrophe(text[index]) &&
                IsWordLetter(text[index - 1]) && IsWordLetter(text[index + 1]);
        }

        private static bool IsApostrophe(char value)
        {
            return value == '\'' || value == '\u02BB' || value == '\u02BC' ||
                value == '\u2018' || value == '\u2019';
        }

        private sealed class WordReplacement
        {
            internal int Start { get; }
            internal int End { get; }
            internal string Text { get; }

            internal WordReplacement(int start, int end, string text)
            {
                Start = start;
                End = end;
                Text = text;
            }
        }

        private static void RunWithBusyIndicator(Action work)
        {
            var app = DocumentHelper.App;
            Cursor previousCursor = Cursor.Current;
            Word.System wordSystem = null;
            Word.WdCursorType previousWordCursor = Word.WdCursorType.wdCursorNormal;
            bool wordCursorChanged = false;
            bool screenUpdating = true;
            bool screenUpdatingChanged = false;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (app != null)
                {
                    try
                    {
                        wordSystem = app.System;
                        previousWordCursor = wordSystem.Cursor;
                        wordSystem.Cursor = Word.WdCursorType.wdCursorWait;
                        wordCursorChanged = true;
                    }
                    catch { }
                    try { app.StatusBar = "Транслитерация бажарилмоқда..."; } catch { }
                    try
                    {
                        screenUpdating = app.ScreenUpdating;
                        app.ScreenUpdating = false;
                        screenUpdatingChanged = true;
                    }
                    catch { }
                }
                work();
            }
            finally
            {
                if (app != null)
                {
                    if (screenUpdatingChanged)
                        try { app.ScreenUpdating = screenUpdating; } catch { }
                    try { app.StatusBar = string.Empty; } catch { }
                    if (screenUpdatingChanged && screenUpdating)
                        try { app.ScreenRefresh(); } catch { }
                }
                if (wordCursorChanged)
                    try { wordSystem.Cursor = previousWordCursor; } catch { }
                if (wordSystem != null && Marshal.IsComObject(wordSystem))
                    try { Marshal.ReleaseComObject(wordSystem); } catch { }
                Cursor.Current = previousCursor;
            }
        }

    }

    public sealed class TransliterationWorkflowResult
    {
        public TransliterationWorkflowStatus Status { get; private set; }
        public string Direction { get; private set; }
        public int WordCount { get; private set; }

        public static TransliterationWorkflowResult NoOp() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.NoOp
            };

        public static TransliterationWorkflowResult ServiceUnavailable() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.ServiceUnavailable
            };

        public static TransliterationWorkflowResult TextNotFound() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.TextNotFound
            };

        public static TransliterationWorkflowResult ScriptUndetermined() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.ScriptUndetermined
            };

        public static TransliterationWorkflowResult Cancelled() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.Cancelled
            };

        public static TransliterationWorkflowResult Completed(string direction, int wordCount) =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.Completed,
                Direction = direction,
                WordCount = Math.Max(0, wordCount)
            };
    }

    public enum TransliterationWorkflowStatus
    {
        NoOp = 0,
        ServiceUnavailable = 1,
        TextNotFound = 2,
        ScriptUndetermined = 3,
        Cancelled = 4,
        Completed = 5
    }
}
