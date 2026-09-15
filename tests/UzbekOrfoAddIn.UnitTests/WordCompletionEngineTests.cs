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
        File.WriteAllText(Path.Combine(directory, "settings.cfg"), "AutoLearn=True\nMaxPredictions=4");
        var settings = new SettingsManager(directory);
        Assert.True(settings.AutoLearn);
        Assert.False(settings.MatnAiLearningConsent);
        Assert.False(settings.PredictionsEnabled);
        settings.MatnAiLearningConsent = true;
        settings.Save();
        var restored = new SettingsManager(directory);
        Assert.True(restored.MatnAiLearningConsent);
        Assert.Equal(4, restored.MaxPredictions);
    }
}
