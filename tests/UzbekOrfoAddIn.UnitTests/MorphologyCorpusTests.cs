using System.Text.Json;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Services;
using Xunit;
using Xunit.Abstractions;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class MorphologyCorpusTests
{
    private readonly ITestOutputHelper _output;

    public MorphologyCorpusTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void EngineeringCorpus_MeetsErrorRateGate()
    {
        string path = Path.Combine(AppContext.BaseDirectory, "Corpora", "uzbek_morphology_corpus.json");
        Corpus corpus = JsonSerializer.Deserialize<Corpus>(File.ReadAllText(path), JsonOptions())!;
        Assert.Equal("uzbekorfo-morphology-corpus-v1", corpus.Schema);
        Assert.Equal("reviewed", corpus.Review.EngineeringStatus);

        var dictionary = new CorpusDictionary(corpus.Lexemes);
        string rulesPath = Path.Combine(AppContext.BaseDirectory, "Data", "uzbek_suffixes.json");
        MorphologyRuleSet rules = JsonSerializer.Deserialize<MorphologyRuleSet>(
            File.ReadAllText(rulesPath), JsonOptions())!;
        Assert.Empty(rules.Validate());
        Assert.Equal(corpus.RuleVersion, rules.Version);
        var parser = new UzbekSuffixParser(dictionary, rules);
        int positives = 0, negatives = 0, falsePositives = 0, falseNegatives = 0;

        foreach (CorpusCase item in corpus.Cases)
        {
            MorphAnalysis analysis = parser.Parse(item.Word);
            bool actual = analysis.IsValidInflectedForm;
            if (actual != item.ExpectedCorrect)
                _output.WriteLine($"MISMATCH {item.Id}: word={item.Word}; expected={item.ExpectedCorrect}; actual={actual}; suffixes={string.Join(",", analysis.SuffixIds)}");
            if (item.ExpectedCorrect)
            {
                positives++;
                if (!actual) falseNegatives++;
                if (actual)
                {
                    Assert.Equal(item.Lemma, analysis.Lemma);
                    Assert.Equal(item.PartOfSpeech, analysis.PartOfSpeech);
                    Assert.Equal(item.SuffixIds, analysis.SuffixIds);
                }
            }
            else
            {
                negatives++;
                if (actual) falsePositives++;
            }
        }

        double falsePositiveRate = negatives == 0 ? 0 : (double)falsePositives / negatives;
        double falseNegativeRate = positives == 0 ? 0 : (double)falseNegatives / positives;
        _output.WriteLine($"positives={positives}; negatives={negatives}; FP={falsePositives} ({falsePositiveRate:P2}); FN={falseNegatives} ({falseNegativeRate:P2})");

        Assert.Equal(0, falsePositives);
        Assert.Equal(0, falseNegatives);
    }

    private static JsonSerializerOptions JsonOptions() => new()
    {
        PropertyNameCaseInsensitive = true
    };

    private sealed class Corpus
    {
        public string Schema { get; set; } = "";
        public string RuleVersion { get; set; } = "";
        public Review Review { get; set; } = new();
        public List<CorpusLexeme> Lexemes { get; set; } = new();
        public List<CorpusCase> Cases { get; set; } = new();
    }

    private sealed class Review
    {
        public string EngineeringStatus { get; set; } = "";
        public string LinguistStatus { get; set; } = "";
    }

    private sealed class CorpusLexeme
    {
        public string Word { get; set; } = "";
        public string Lemma { get; set; } = "";
        public string PartOfSpeech { get; set; } = "";
    }

    private sealed class CorpusCase
    {
        public string Id { get; set; } = "";
        public string Word { get; set; } = "";
        public bool ExpectedCorrect { get; set; }
        public string? Lemma { get; set; }
        public string? PartOfSpeech { get; set; }
        public List<string>? SuffixIds { get; set; }
    }

    private sealed class CorpusDictionary : IDictionaryService, ILexemeMetadataProvider
    {
        private readonly Dictionary<string, LexemeMetadata> _entries;

        public CorpusDictionary(IEnumerable<CorpusLexeme> entries)
        {
            _entries = entries.ToDictionary(
                e => e.Word,
                e => new LexemeMetadata
                {
                    Word = e.Word,
                    Lemma = e.Lemma,
                    PartOfSpeech = e.PartOfSpeech
                },
                StringComparer.OrdinalIgnoreCase);
        }

        public bool Contains(string word) => _entries.ContainsKey(word);
        public bool TryGetLexeme(string word, out LexemeMetadata lexeme) =>
            _entries.TryGetValue(word, out lexeme!);
        public bool IsMainDictionaryWord(string word) => Contains(word);
        public int TotalWordCount => _entries.Count;
        public void Load() { }
        public void Save() { }
        public AddWordResult AddWord(string word) => AddWordResult.AlreadyInUserDictionary;
        public void RemoveWord(string word) { }
        public List<string> GetUserWords() => new();
        public List<string> GetMainWords() => _entries.Keys.ToList();
        public List<string> GetAllWords() => _entries.Keys.ToList();
        public int ImportFromFile(string filePath) => 0;
        public int ExportForMigration(string filePath) => 0;
    }
}
