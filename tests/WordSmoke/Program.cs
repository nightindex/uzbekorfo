using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
        string temporaryDirectory = null;
        try
        {
            app = new Word.Application();
            // Never close or alter an application that already has a user's documents.
            if (app.Documents.Count != 0) throw new InvalidOperationException("Expected a new empty Word instance.");
            ownsApplication = true;
            Console.WriteLine("ENVIRONMENT: Word version=" + app.Version + "; build=" + app.Build +
                "; harness=" + (Environment.Is64BitProcess ? "x64" : "x86") +
                ". Harness bitness is not proof of Office bitness. This is not an installation test.");
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

            string dataDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            temporaryDirectory = Path.Combine(Path.GetTempPath(), "UzbekOrfoWordSmoke", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temporaryDirectory);
            var dictionary = new DictionaryService(
                Path.Combine(dataDirectory, "uzbek_main.dic"),
                Path.Combine(temporaryDirectory, "user_custom.dic"),
                Path.Combine(dataDirectory, "uzbek_dictionary_metadata.json"));
            dictionary.Load();
            var transliterator = new TransliterationService(Path.Combine(dataDirectory, "translit_exceptions.json"));
            dictionary.SetTransliterator(transliterator);
            var morphology = new UzbekMorphAnalyzer(dictionary, transliterator);
            morphology.LoadSuffixes(Path.Combine(dataDirectory, "uzbek_suffixes.json"));
            Require(morphology.LoadedSuffixCount >= 50, "Versioned morphology rules loaded");
            var loadedRules = MorphologyRuleSetLoader.Load(Path.Combine(dataDirectory, "uzbek_suffixes.json"));
            Require(loadedRules.Suffixes.Count(s => s.StandaloneOnly && !string.IsNullOrEmpty(s.RequiredRootFlags)) == 2333,
                "Runtime JSON loader preserves imported rule restrictions");
            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            string cleanSuffixPath = Path.Combine(temporaryDirectory, "clean-suffixes.json");
            DataSeedService.SeedGrammarFiles(cleanSuffixPath, Path.Combine(temporaryDirectory, "clean-grammar.json"),
                Path.Combine(temporaryDirectory, "clean-proper.json"));
            Require(MorphologyRuleSetLoader.Load(cleanSuffixPath).Suffixes.Count == loadedRules.Suffixes.Count,
                "Fresh data seeding includes all bundled suffixes (not a clean installer test)");
            var legacy = serializer.Deserialize<Dictionary<string, object>>(File.ReadAllText(Path.Combine(dataDirectory, "uzbek_suffixes.json")));
            foreach (string section in new[] { "suffixes", "families" })
                legacy[section] = ((System.Collections.IEnumerable)legacy[section]).Cast<Dictionary<string, object>>()
                    .Where(row => !((string)row["id"]).StartsWith("IMPORTED_", StringComparison.Ordinal) && (string)row["id"] != "imported").ToArray();
            legacy["localNote"] = "preserve-me";
            string upgradePath = Path.Combine(temporaryDirectory, "suffixes-upgrade.json");
            File.WriteAllText(upgradePath, serializer.Serialize(legacy));
            DataSeedService.SeedGrammarFiles(upgradePath, Path.Combine(temporaryDirectory, "grammar.json"), Path.Combine(temporaryDirectory, "proper.json"));
            Require(MorphologyRuleSetLoader.Load(upgradePath).Suffixes.Count == loadedRules.Suffixes.Count &&
                File.ReadAllText(upgradePath).Contains("preserve-me") && File.Exists(upgradePath + ".bak"),
                "Existing suffix file receives compatible rules while retaining custom data and backup");
            string upgraded = File.ReadAllText(upgradePath);
            DataSeedService.SeedGrammarFiles(upgradePath, Path.Combine(temporaryDirectory, "grammar.json"), Path.Combine(temporaryDirectory, "proper.json"));
            Require(File.ReadAllText(upgradePath) == upgraded, "Suffix upgrade is idempotent");
            LexemeMetadata kitobLexeme;
            Require(dictionary.TryGetLexeme("китоб", out kitobLexeme) &&
                kitobLexeme.Lemma == "китоб" && kitobLexeme.PartOfSpeech == "noun",
                "Bundled lemma and part-of-speech metadata loaded");
            Require(morphology.Analyze("китобларимиздан").IsValidInflectedForm,
                "Productive suffix chain works against bundled data");
            Require(morphology.Analyze("kitoblarimizdan").IsValidInflectedForm,
                "Productive suffix chain works in Latin script");
            Require(!morphology.Analyze("китобаман").IsValidInflectedForm,
                "Part-of-speech metadata blocks a noun with a verb chain");
            Require(!morphology.Analyze("бормоқлар").IsValidInflectedForm,
                "Infinitive cannot be followed by finite agreement");
            Require(morphology.Analyze("отам").IsValidInflectedForm &&
                !morphology.Analyze("отаим").IsValidInflectedForm,
                "Vowel-final possessive allomorph is enforced");
            Require(morphology.Analyze("мактабда").IsValidInflectedForm &&
                !morphology.Analyze("мактабта").IsValidInflectedForm,
                "Written locative uses invariant -да");
            Require(morphology.Analyze("дарахтдан").IsValidInflectedForm &&
                !morphology.Analyze("дарахттан").IsValidInflectedForm,
                "Written ablative uses invariant -дан");

            var unknownWords = new UnknownWordReportService(
                Path.Combine(temporaryDirectory, "unknown_words.tsv"));
            var spelling = new SpellingEngine(dictionary, morphology, unknownWords);
            a.Content.Text = "китоп китоб";
            Require(spelling.Check(a.Content).Count == 1, "Unknown word is detected in Word");
            Require(unknownWords.GetSnapshot().Count == 1 &&
                unknownWords.GetSnapshot()[0].Word == "китоп",
                "Unknown word is aggregated into the local review report");
            const int largeDocumentWordCount = 50000;
            a.Content.Text = "kitoblarimizdan китобларимиздан";
            var importedErrors = spelling.Check(a.Content);
            Require(importedErrors.Count == 0, "Existing morphology engine handles both scripts");
            Require(morphology.Analyze("китоблар-а").IsValidInflectedForm,
                "Compatible imported suffix uses the existing parser and root flags");
            Require(spelling.IsCorrect("kitoblar-a"), "Imported complete suffix works in Latin");
            Require(!spelling.IsCorrect("kitoblar-alar"), "Imported complete suffix cannot be chained");
            dictionary.RebuildSearchIndexBackground();
            Require(spelling.GetSuggestions("kitb", 5).Any(), "Existing suggestions are available");
            a.Content.Text = string.Join(" ", Enumerable.Repeat("китоб", largeDocumentWordCount));
            var stopwatch = Stopwatch.StartNew();
            List<ErrorEntry> largeDocumentErrors = spelling.Check(a.Content);
            stopwatch.Stop();
            Require(largeDocumentErrors.Count == 0, "Large known-word document has no spelling errors");
            Require(stopwatch.Elapsed < TimeSpan.FromSeconds(30),
                "Large-document check stays under the 30-second release ceiling; actual=" + stopwatch.Elapsed);
            Console.WriteLine("PERF: " + largeDocumentWordCount + " Word tokens checked in " +
                stopwatch.Elapsed.TotalMilliseconds.ToString("F0") + " ms");
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
            if (!string.IsNullOrWhiteSpace(temporaryDirectory))
                try { Directory.Delete(temporaryDirectory, true); } catch { }
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
