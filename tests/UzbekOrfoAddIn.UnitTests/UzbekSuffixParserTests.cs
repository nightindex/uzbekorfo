using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Services;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class UzbekSuffixParserTests
{
    private readonly FakeDictionary _dictionary = new(
        ("китоб", "noun"), ("бормоқ", "verb"), ("юрак", "noun"), ("ота", "noun"));

    [Fact]
    public void Parse_AcceptsPluralPossessiveCaseChain()
    {
        var result = CreateParser().Parse("китобларимиздан");

        Assert.True(result.IsValidInflectedForm);
        Assert.Equal("китоб", result.Root);
        Assert.Equal("китоб", result.Lemma);
        Assert.Equal("noun", result.PartOfSpeech);
        Assert.Equal(new[] { "PL", "POSS_1P", "ABL_DAN" }, result.SuffixIds);
    }

    [Fact]
    public void Parse_RejectsInvalidNounSuffixOrder()
    {
        var result = CreateParser().Parse("китобимизлардан");

        Assert.True(result.IsKnownRoot);
        Assert.False(result.IsValidInflectedForm);
    }

    [Fact]
    public void Parse_RejectsUnknownRootEvenWhenSuffixesLookValid()
    {
        var result = CreateParser().Parse("зотларимиздан");

        Assert.False(result.IsKnownRoot);
        Assert.False(result.IsValidInflectedForm);
    }

    [Fact]
    public void Parse_RequiresVerbInfinitiveAndCompatiblePartOfSpeech()
    {
        Assert.True(CreateParser().Parse("бораман").IsValidInflectedForm);
        Assert.False(CreateParser().Parse("китобаман").IsValidInflectedForm);
        Assert.False(CreateParser().Parse("бормоқлар").IsValidInflectedForm);
    }

    [Fact]
    public void Parse_ValidatesConditionedAllomorphs()
    {
        Assert.True(CreateParser().Parse("юракка").IsValidInflectedForm);
        Assert.False(CreateParser().Parse("юракга").IsValidInflectedForm);
    }

    [Fact]
    public void Parse_ValidatesPossessiveRootMutation()
    {
        Assert.True(CreateParser().Parse("юрагим").IsValidInflectedForm);
        Assert.False(CreateParser().Parse("юраким").IsValidInflectedForm);
    }

    [Fact]
    public void Parse_UsesVowelPossessiveAllomorph()
    {
        Assert.True(CreateParser().Parse("отаси").IsValidInflectedForm);
        Assert.False(CreateParser().Parse("отаи").IsValidInflectedForm);
    }

    [Fact]
    public void RuleSetValidation_RejectsBrokenReferences()
    {
        MorphologyRuleSet rules = Rules();
        rules.AllomorphGroups[0].SuffixIds.Add("MISSING");

        Assert.Contains(rules.Validate(), error => error.Contains("unknown suffix"));
        Assert.Throws<ArgumentException>(() => new UzbekSuffixParser(_dictionary, rules));
    }

    private UzbekSuffixParser CreateParser() => new(_dictionary, Rules());

    [Theory]
    [InlineData("A", "китоблар-а", true)]
    [InlineData("a", "китоблар-а", false)]
    [InlineData("", "китоблар-а", false)]
    [InlineData("A", "китоблар-алар", false)]
    [InlineData("A", "зотлар-а", false)]
    public void ImportedCompleteSuffixRequiresFlagAndCannotBeChained(string flags, string word, bool expected)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Data", "uzbek_suffixes.json");
        var rules = System.Text.Json.JsonSerializer.Deserialize<MorphologyRuleSet>(File.ReadAllText(path),
            new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;
        Assert.Empty(rules.Validate());
        _dictionary.TryGetLexeme("китоб", out var lexeme);
        lexeme.HunspellFlags = flags;
        Assert.Equal(expected, new UzbekSuffixParser(_dictionary, rules).Parse(word).IsValidInflectedForm);
    }

    internal static MorphologyRuleSet Rules()
    {
        return new MorphologyRuleSet
        {
            Version = "uzbekorfo-suffixes-v2",
            MaxSuffixDepth = 6,
            MinimumRootLength = 2,
            Vowels = new() { "а", "е", "ё", "и", "о", "у", "ў", "э", "ю", "я" },
            RootMutations = new()
            {
                new RootMutationRule
                {
                    Id = "voice",
                    FinalCharacterMap = new() { ["к"] = "г", ["қ"] = "ғ" }
                }
            },
            Families = new()
            {
                new MorphologyFamilyRule
                {
                    Id = "noun",
                    RootLookupSuffix = "",
                    AllowedRootPartOfSpeech = new() { "noun", "adjective" },
                    AllowUnknownRootPartOfSpeech = true,
                    CategoryStages = new()
                    {
                        ["plural"] = 1,
                        ["possessive"] = 2,
                        ["case_dative"] = 3,
                        ["case_ablative"] = 3
                    },
                    OccurrenceLimits = new()
                    {
                        new CategoryOccurrenceLimit { Categories = new() { "plural" }, Max = 1 },
                        new CategoryOccurrenceLimit { Categories = new() { "possessive" }, Max = 1 },
                        new CategoryOccurrenceLimit
                        {
                            Categories = new() { "case_dative", "case_ablative" }, Max = 1
                        }
                    }
                },
                new MorphologyFamilyRule
                {
                    Id = "verb",
                    RootLookupSuffix = "моқ",
                    AllowedRootPartOfSpeech = new() { "verb" },
                    AllowUnknownRootPartOfSpeech = true,
                    CategoryStages = new()
                    {
                        ["verb_tense"] = 1,
                        ["verb_agreement"] = 2
                    },
                    OccurrenceLimits = new()
                    {
                        new CategoryOccurrenceLimit { Categories = new() { "verb_tense" }, Max = 1 },
                        new CategoryOccurrenceLimit { Categories = new() { "verb_agreement" }, Max = 1 }
                    },
                    Requirements = new()
                    {
                        new CategoryRequirementRule
                        {
                            IfAny = new() { "verb_agreement" },
                            RequireAny = new() { "verb_tense" }
                        }
                    }
                }
            },
            Suffixes = new()
            {
                Suffix("PL", "лар", "plural", "noun"),
                Suffix("POSS_1P", "имиз", "possessive", "noun", mutation: "voice"),
                Suffix("POSS_1S", "им", "possessive", "noun", mutation: "voice"),
                Suffix("POSS_3S", "и", "possessive", "noun", mutation: "voice"),
                Suffix("POSS_3S_SI", "си", "possessive", "noun", "stemEndsInVowel", "voice"),
                Suffix("ABL_DAN", "дан", "case_ablative", "noun", "default"),
                Suffix("DAT_GA", "га", "case_dative", "noun", "default"),
                Suffix("DAT_KA", "ка", "case_dative", "noun", "stemEndsIn:к"),
                Suffix("V_PRES", "а", "verb_tense", "verb"),
                Suffix("V_PRES_Y", "й", "verb_tense", "verb", "stemEndsInVowel"),
                Suffix("AGR_MAN", "ман", "verb_agreement", "verb")
            },
            AllomorphGroups = new()
            {
                new AllomorphGroupRule
                {
                    Id = "dative", SuffixIds = new() { "DAT_KA", "DAT_GA" }, FallbackSuffixId = "DAT_GA"
                },
                new AllomorphGroupRule
                {
                    Id = "possessiveThirdSingular",
                    SuffixIds = new() { "POSS_3S_SI", "POSS_3S" },
                    FallbackSuffixId = "POSS_3S"
                },
                new AllomorphGroupRule
                {
                    Id = "presentTense",
                    SuffixIds = new() { "V_PRES_Y", "V_PRES" },
                    FallbackSuffixId = "V_PRES"
                }
            }
        };
    }

    private static SuffixEntry Suffix(
        string id,
        string cyrillic,
        string category,
        string family,
        string? condition = null,
        string? mutation = null)
    {
        return new SuffixEntry
        {
            Id = id,
            Cyrillic = cyrillic,
            Latin = cyrillic,
            Category = category,
            Family = family,
            Productive = true,
            AllomorphCondition = condition,
            RootMutation = mutation
        };
    }

    private sealed class FakeDictionary : IDictionaryService, ILexemeMetadataProvider
    {
        private readonly HashSet<string> _words;
        private readonly Dictionary<string, LexemeMetadata> _metadata;

        public FakeDictionary(params (string Word, string PartOfSpeech)[] entries)
        {
            _words = entries.Select(e => e.Word).ToHashSet(StringComparer.OrdinalIgnoreCase);
            _metadata = entries.ToDictionary(
                e => e.Word,
                e => new LexemeMetadata { Word = e.Word, Lemma = e.Word, PartOfSpeech = e.PartOfSpeech },
                StringComparer.OrdinalIgnoreCase);
        }

        public bool TryGetLexeme(string word, out LexemeMetadata lexeme) =>
            _metadata.TryGetValue(word, out lexeme!);
        public bool Contains(string word) => _words.Contains(word);
        public bool IsMainDictionaryWord(string word) => Contains(word);
        public int TotalWordCount => _words.Count;
        public void Load() { }
        public void Save() { }
        public AddWordResult AddWord(string word) =>
            _words.Add(word) ? AddWordResult.Added : AddWordResult.AlreadyInUserDictionary;
        public void RemoveWord(string word) => _words.Remove(word);
        public List<string> GetUserWords() => new();
        public List<string> GetMainWords() => _words.ToList();
        public List<string> GetAllWords() => _words.ToList();
        public int ImportFromFile(string filePath) => 0;
        public int ExportForMigration(string filePath) => 0;
    }
}
