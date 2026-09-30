using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using UzbekOrfoAddIn.Services;

// Read-only evaluation against the production dictionary and spelling engine, without Word.
internal static class Program
{
    private static string Hash(string path)
    {
        using (var stream = File.OpenRead(path))
        using (var sha = SHA256.Create())
            return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "");
    }

    private static int Main(string[] args)
    {
        try
        {
            if (args.Length != 3) throw new ArgumentException("Expected data directory, corpus path, report path.");
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var corpus = json.Deserialize<Corpus>(File.ReadAllText(args[1]));
            if (corpus.schema != "uzbekorfo-spelling-evaluation-v1" || corpus.cases == null || corpus.cases.Length == 0)
                throw new InvalidDataException("Provide a spelling evaluation corpus, not the four-root parser fixture.");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in corpus.cases)
                if (string.IsNullOrWhiteSpace(item.id) || !ids.Add(item.id) || string.IsNullOrWhiteSpace(item.word) ||
                    !item.expectedCorrect.HasValue || string.IsNullOrWhiteSpace(item.source) || string.IsNullOrWhiteSpace(item.category))
                    throw new InvalidDataException("Each case requires a unique id, word, boolean expectedCorrect, source and category.");
            var dictionary = new DictionaryService(Path.Combine(args[0], "uzbek_main.dic"),
                Path.Combine(Path.GetTempPath(), "uzbekorfo-evaluation-" + Guid.NewGuid().ToString("N"), "user.dic"),
                Path.Combine(args[0], "uzbek_dictionary_metadata.json"));
            dictionary.Load();
            var transliterator = new TransliterationService(Path.Combine(args[0], "translit_exceptions.json"));
            dictionary.SetTransliterator(transliterator);
            var morphology = new UzbekMorphAnalyzer(dictionary, transliterator);
            morphology.LoadSuffixes(Path.Combine(args[0], "uzbek_suffixes.json"));
            var spelling = new SpellingEngine(dictionary, morphology);
            var rows = corpus.cases.Select(c => new {
                c.id, c.word, c.category, c.source, expectedCorrect = c.expectedCorrect.Value,
                actualCorrect = spelling.IsCorrect(c.word)
            }).ToArray();
            int valid = rows.Count(r => r.expectedCorrect), invalid = rows.Length - valid;
            if (valid == 0 || invalid == 0) throw new InvalidDataException("Both valid and invalid labelled forms are required.");
            int falseAlarms = rows.Count(r => r.expectedCorrect && !r.actualCorrect);
            int missedErrors = rows.Count(r => !r.expectedCorrect && r.actualCorrect);
            // Positive means an error was flagged, not that a word was accepted.
            var report = new {
                schema = "uzbekorfo-spelling-results-v1", generatedAtUtc = DateTime.UtcNow.ToString("o"),
                scope = "Engineering measurement only; independent review is a separate release gate.",
                corpusSha256 = Hash(args[1]), engineSha256 = Hash(typeof(SpellingEngine).Assembly.Location),
                dataHashes = new[] { "uzbek_main.dic", "uzbek_dictionary_metadata.json", "uzbek_suffixes.json", "translit_exceptions.json" }
                    .ToDictionary(f => f, f => Hash(Path.Combine(args[0], f))),
                valid, invalid, truePositives = invalid - missedErrors, trueNegatives = valid - falseAlarms,
                falsePositives = falseAlarms, falseNegatives = missedErrors,
                falsePositiveRate = (double)falseAlarms / valid, falseNegativeRate = (double)missedErrors / invalid,
                byCategory = rows.GroupBy(r => r.category).Select(g => new {
                    category = g.Key, valid = g.Count(r => r.expectedCorrect), invalid = g.Count(r => !r.expectedCorrect),
                    falseAlarms = g.Count(r => r.expectedCorrect && !r.actualCorrect),
                    missedErrors = g.Count(r => !r.expectedCorrect && r.actualCorrect)
                }).ToArray(), cases = rows
            };
            File.WriteAllText(args[2], json.Serialize(report));
            Console.WriteLine("ENGINEERING ONLY: false alarms={0}/{1}; missed errors={2}/{3}. Report: {4}",
                falseAlarms, valid, missedErrors, invalid, args[2]);
            return 0; // Measurement, not an accuracy acceptance threshold.
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.Message); return 1; }
    }

    public sealed class Corpus { public string schema { get; set; } public Case[] cases { get; set; } }
    public sealed class Case
    {
        public string id { get; set; } public string word { get; set; }
        public bool? expectedCorrect { get; set; } public string source { get; set; } public string category { get; set; }
    }
}
