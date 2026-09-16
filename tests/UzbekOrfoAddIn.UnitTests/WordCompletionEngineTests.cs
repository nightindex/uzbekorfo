using UzbekOrfoAddIn.Services;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public class WordCompletionEngineTests
{
    [Fact]
    public void CompletionPreservesTypedCaseAndApostrophe()
    {
        var engine = new WordCompletionEngine(new[] { "o'zbek", "kitob", "kitoblar", "китоб", "китоблар" });
        Assert.Equal(new[] { "O‘zbek" }, engine.Complete("O‘z"));
        Assert.Equal(new[] { "KITOB", "KITOBLAR" }, engine.Complete("KIT"));
        Assert.Equal(new[] { "китоб", "китоблар" }, engine.Complete("кит"));
    }
    [Fact]
    public void ExcludesMixedScriptsPhrasesAndExactWords()
    {
        var engine = new WordCompletionEngine(new[] { "kitob", "kitob", "kitоб", "kitob yaxshi", "kitoblar" });
        Assert.Equal(new[] { "kitoblar" }, engine.Complete("kitob"));
        Assert.Empty(engine.Complete("k"));
        Assert.Empty(engine.Complete("kit", 0));
    }
    [Fact]
    public void SuffixCandidateRequiresValidatorAndPartialEnding()
    {
        var engine = new WordCompletionEngine(new[] { "kitob" }, new[] { "lar", "larlar" }, w => w == "kitoblar");
        Assert.Equal(new[] { "kitoblar" }, engine.Complete("kitobl"));
        Assert.Empty(new WordCompletionEngine(new[] { "kitob" }, new[] { "lar" }).Complete("kitobl"));
    }
    [Fact]
    public void CancellationAndLimitAreRespected()
    {
        var engine = new WordCompletionEngine(new[] { "kitob", "kitoblar", "kitobcha" });
        Assert.Single(engine.Complete("kit", 1));
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => engine.Complete("kit", 3, cancel.Token));
        Assert.Throws<OperationCanceledException>(() => new WordCompletionEngine(new[] { "kitob" }, cancellation: cancel.Token));
    }

    [Fact]
    public void DeferredProposalsAreExplicitlyUnvalidated()
    {
        var engine = new WordCompletionEngine(new[] { "kitob" }, new[] { "larlar" });
        Assert.Empty(engine.Complete("kitobl"));
        Assert.Equal(new[] { "kitoblarlar" }, engine.Complete("kitobl", 3, default, deferMorphologyValidation: true));
        Assert.Empty(WordCompletionEngine.ValidateCandidates(engine.Complete("kitobl", 3, default, true), _ => false, 3));
    }

    [Fact]
    public void PersonalCountsAreOptionalAndRoundTrip()
    {
        string directory = Path.Combine(Path.GetTempPath(), "MatnAiTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var preferences = new CompletionPreferences(Path.Combine(directory, "counts.tsv"));
        preferences.Record("kitoblar");
        preferences.Record("not a token");
        preferences.Save();
        var restored = new CompletionPreferences(Path.Combine(directory, "counts.tsv"));
        restored.Load();
        Assert.Single(restored.Snapshot());
        var engine = new WordCompletionEngine(new[] { "kitob", "kitoblar" });
        Assert.Equal("kitob", engine.Complete("kit")[0]);
        Assert.Equal("kitoblar", engine.Complete("kit", acceptanceCounts: restored.Snapshot())[0]);
        restored.Clear();
        restored.Load();
        Assert.Empty(restored.Snapshot());
    }

    [Fact]
    public void PersonalOverlayRetrievesBeyondLexicalCandidateWindow()
    {
        var words = Enumerable.Range(0, 100).Select(i => "ki" + (char)('a' + i / 26) + (char)('a' + i % 26)).Append("kiz").ToArray();
        var engine = new WordCompletionEngine(words);
        Assert.DoesNotContain("kiz", engine.Complete("ki"));
        Assert.Equal("kiz", engine.Complete("ki", acceptanceCounts: new Dictionary<string, int> { ["kiz"] = 2 })[0]);
    }

    [Fact]
    public void LegacyAutoLearnDoesNotImplyMatnAiConsent()
    {
        string directory = Path.Combine(Path.GetTempPath(), "MatnAiTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, "settings.cfg"),
            "AutoLearn=True\nMaxPredictions=4\nMinPredictionLength=2");
        var settings = new SettingsManager(directory);
        Assert.True(settings.AutoLearn);
        Assert.False(settings.MatnAiLearningConsent);
        Assert.False(settings.PredictionsEnabled);
        Assert.Equal(2, settings.MinPredictionLength);
        settings.MatnAiLearningConsent = true;
        settings.Save();
        var restored = new SettingsManager(directory);
        Assert.True(restored.MatnAiLearningConsent);
        Assert.Equal(4, restored.MaxPredictions);
    }

    [Theory]
    [InlineData("ki", "kitob", "tob")]
    [InlineData("КИ", "КИТОБ", "ТОБ")]
    [InlineData("k", "kitob", null)]
    [InlineData("so'", "so'fi", "fi")]
    [InlineData("kit", "kita", "a")]
    [InlineData("kit", "other", null)]
    public void GhostTailRequiresUsefulExactContinuation(string prefix, string candidate, string? expected)
    {
        Assert.Equal(expected, WordCompletionEngine.GetGhostTail(prefix, candidate));
    }

    [Fact]
    public void GhostTailHonorsConfiguredMinimum()
    {
        Assert.Null(WordCompletionEngine.GetGhostTail("ki", "kitob", 3));
        Assert.Equal("ob", WordCompletionEngine.GetGhostTail("kit", "kitob", 3));
    }

    [Theory]
    [InlineData("Sal", "Sala", "Salom", "Salomat")]
    [InlineData("Сал", "Сала", "Салом", "Саломат")]
    public void WordCompletionPrefersUsefulTailAndStaysStableWhileTyping(
        string prefix, string shortWord, string wholeWord, string longerWord)
    {
        var engine = new WordCompletionEngine(new[] { shortWord, wholeWord, longerWord });
        Assert.Equal(wholeWord, engine.Complete(prefix, 1)[0]);
        string nextPrefix = wholeWord.Substring(0, wholeWord.Length - 1);
        Assert.Equal(wholeWord, engine.Complete(nextPrefix, 1, continuedWord: wholeWord)[0]);
        Assert.Single(WordCompletionEngine.GetGhostTail(nextPrefix, wholeWord)!);
        Assert.Equal(longerWord, engine.Complete(wholeWord, 1, continuedWord: wholeWord)[0]);
        Assert.Empty(engine.Complete("xyz", continuedWord: wholeWord));
    }

    [Fact]
    public void RetainedWordMustStillBelongToVocabulary()
    {
        var engine = new WordCompletionEngine(new[] { "kitob" });
        Assert.Equal(new[] { "kitob" }, engine.Complete("ki", continuedWord: "kitxyz"));
    }
}
