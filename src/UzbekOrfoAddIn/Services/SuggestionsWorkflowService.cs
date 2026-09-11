using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.UI;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates suggestions flow for the currently selected word:
    /// grammar-aware suggestions first, then spelling suggestions.
    /// </summary>
    public sealed class SuggestionsWorkflowService
    {
        private readonly SpellingEngine _spellingEngine;
        private readonly GrammarEngine _grammarEngine;
        private readonly Word.Application _wordApplication;
        private readonly Func<string, AddWordResult> _addWordToDictionary;
        private readonly Action<string, string> _replaceAllInDocument;

        public SuggestionsWorkflowService(
            SpellingEngine spellingEngine,
            GrammarEngine grammarEngine,
            Word.Application wordApplication,
            Func<string, AddWordResult> addWordToDictionary,
            Action<string, string> replaceAllInDocument)
        {
            _spellingEngine = spellingEngine;
            _grammarEngine = grammarEngine;
            _wordApplication = wordApplication;
            _addWordToDictionary = addWordToDictionary;
            _replaceAllInDocument = replaceAllInDocument;
        }

        public void Execute(Image titleIcon)
        {
            if (_spellingEngine == null)
            {
                SafeExecutor.ShowWarning("Имло текширув тизими юкланмаган.");
                return;
            }

            string selectedText = DocumentHelper.GetSelectedWord();
            if (string.IsNullOrWhiteSpace(selectedText))
            {
                SafeExecutor.ShowInfo("Аввал сўзни танланг.");
                return;
            }

            string word = selectedText.Trim();
            Word.Selection selection = _wordApplication?.Selection;
            Word.Range replaceRange = ResolveReplaceRange(selection);

            var grammarMatch = FindGrammarMatch(word, selection, out Word.Range grammarReplaceRange);
            var spellingSuggestions = _spellingEngine.GetSuggestions(word);

            if (grammarMatch != null)
            {
                ShowGrammarAwareSuggestionsForm(
                    word,
                    spellingSuggestions,
                    grammarMatch,
                    grammarReplaceRange,
                    replaceRange,
                    selection,
                    titleIcon);
                return;
            }

            if (spellingSuggestions.Count == 0)
            {
                ModernMessageBox.Show(
                    $"\"{word}\" учун вариант топилмади.",
                    "Вариантлар",
                    ModernMessageBox.MessageType.Info,
                    "Тушундим",
                    titleIcon);
                return;
            }

            ShowSpellingSuggestionsForm(word, spellingSuggestions, replaceRange, selection, titleIcon);
        }

        private static Word.Range ResolveReplaceRange(Word.Selection selection)
        {
            try
            {
                if (selection == null || selection.Range == null) return null;

                string currentSelection = selection.Range.Text ?? string.Empty;
                if (selection.Type == Word.WdSelectionType.wdSelectionNormal &&
                    !string.IsNullOrWhiteSpace(currentSelection) &&
                    currentSelection.Trim().Length > 0)
                {
                    return selection.Range.Duplicate;
                }

                var wordRange = selection.Range.Words[1];
                if (wordRange != null && !string.IsNullOrWhiteSpace(wordRange.Text))
                    return wordRange.Duplicate;
            }
            catch { }

            return null;
        }

        private ErrorEntry FindGrammarMatch(string word, Word.Selection selection, out Word.Range grammarReplaceRange)
        {
            grammarReplaceRange = null;
            if (_grammarEngine == null || selection == null || selection.Range == null)
                return null;

            try
            {
                var sentenceRange = selection.Range.Sentences[1];
                string sentenceText = sentenceRange?.Text ?? string.Empty;
                if (string.IsNullOrEmpty(sentenceText))
                    return null;

                var grammarErrors = _grammarEngine.CheckText(sentenceText);
                int wordStart = selection.Range.Start - sentenceRange.Start;
                int wordEnd = wordStart + word.Length;

                var grammarMatch = grammarErrors.FirstOrDefault(ge =>
                    ge.StartIndex < wordEnd && ge.EndIndex > wordStart);

                if (grammarMatch != null)
                {
                    int absStart = sentenceRange.Start + Math.Max(0, grammarMatch.StartIndex);
                    int absEnd = sentenceRange.Start + Math.Max(grammarMatch.EndIndex, grammarMatch.StartIndex + 1);
                    if (absEnd > absStart && absEnd <= sentenceRange.Document.Content.End)
                        grammarReplaceRange = sentenceRange.Document.Range(absStart, absEnd);
                }

                return grammarMatch;
            }
            catch
            {
                return null;
            }
        }

        private static List<Suggestion> BuildGrammarAwareSuggestions(
            ErrorEntry grammarMatch,
            List<Suggestion> spellingSuggestions)
        {
            var merged = new List<Suggestion>();
            if (!string.IsNullOrEmpty(grammarMatch?.BestSuggestion))
            {
                merged.Add(new Suggestion
                {
                    Text = grammarMatch.BestSuggestion,
                    Confidence = 1.0
                });
            }

            foreach (var suggestion in spellingSuggestions ?? new List<Suggestion>())
            {
                if (merged.All(existing => existing.Text != suggestion.Text))
                    merged.Add(suggestion);
            }

            return merged;
        }

        private void ShowGrammarAwareSuggestionsForm(
            string word,
            List<Suggestion> spellingSuggestions,
            ErrorEntry grammarMatch,
            Word.Range grammarReplaceRange,
            Word.Range replaceRange,
            Word.Selection selection,
            Image titleIcon)
        {
            var suggestions = BuildGrammarAwareSuggestions(grammarMatch, spellingSuggestions);
            var target = CaptureTarget(grammarReplaceRange ?? replaceRange, trimTrailingWhitespace: false);

            var form = new SuggestionsForm(
                word,
                suggestions,
                onReplace: replacement =>
                {
                    SafeExecutor.Execute(() =>
                    {
                        ApplySuggestion(target, replacement);
                    }, "Алмаштириш");
                },
                onReplaceAll: null,
                onAddToDict: null,
                isGrammarError: true,
                grammarMessage: grammarMatch?.Message);

            try { form.TitleIcon = titleIcon; } catch { }
            form.FormClosed += (s, e) => ReleaseTarget(target);
            form.Show();
        }

        private void ShowSpellingSuggestionsForm(
            string word,
            List<Suggestion> suggestions,
            Word.Range replaceRange,
            Word.Selection selection,
            Image titleIcon)
        {
            var target = CaptureTarget(replaceRange, trimTrailingWhitespace: true);
            var form = new SuggestionsForm(
                word,
                suggestions,
                onReplace: replacement =>
                {
                    SafeExecutor.Execute(() =>
                    {
                        ApplySuggestion(target, replacement);
                    }, "Алмаштириш");
                },
                onReplaceAll: replacement =>
                {
                    SafeExecutor.Execute(() =>
                    {
                        if (target?.Range == null ||
                            !DocumentHelper.IsSameDocument(target.Range.Document, DocumentHelper.ActiveDoc))
                        {
                            SafeExecutor.ShowWarning("Ҳужжат ўзгарган. Керакли ҳужжатда вариантларни қайта очинг.");
                            return;
                        }
                        _replaceAllInDocument?.Invoke(word, replacement);
                    }, "Барчасини алмаштириш");
                },
                onAddToDict: () =>
                {
                    if (_addWordToDictionary == null)
                        return AddWordResult.Invalid;
                    return _addWordToDictionary(word);
                });

            try { form.TitleIcon = titleIcon; } catch { }
            form.FormClosed += (s, e) => ReleaseTarget(target);
            form.Show();
        }

        private static ErrorEntry CaptureTarget(Word.Range range, bool trimTrailingWhitespace)
        {
            if (range == null) return null;
            var owned = range.Duplicate;
            if (trimTrailingWhitespace)
                owned.End = owned.Start + (owned.Text ?? string.Empty).TrimEnd().Length;
            return new ErrorEntry { Range = owned, Word = owned.Text, OriginalText = owned.Text };
        }

        private static void ApplySuggestion(ErrorEntry target, string replacement)
        {
            DocumentHelper.BeginUndoRecord("Алмаштириш");
            try
            {
                if (!DocumentHelper.TryReplaceError(target, DocumentHelper.ActiveDoc, replacement))
                    SafeExecutor.ShowWarning("Ҳужжат ёки матн ўзгарган. Вариантларни қайта очинг.");
            }
            finally { DocumentHelper.EndUndoRecord(); }
        }

        private static void ReleaseTarget(ErrorEntry target)
        {
            if (target?.Range == null) return;
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(target.Range); } catch { }
            target.Range = null;
        }
    }
}
