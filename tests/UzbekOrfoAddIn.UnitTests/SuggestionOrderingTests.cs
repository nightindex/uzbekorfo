using UzbekOrfoAddIn.Prediction;
using UzbekOrfoAddIn.Models;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class SuggestionOrderingTests
{
    [Fact]
    public void CurrentPolicyPreservesPhraseFirstOrdering()
    {
        var weakPhrase = Candidate("qaror", 10);
        Assert.Equal(new[] { "qaror", "qalam" }, SuggestionOrdering.Merge("qa",
            new[] { weakPhrase }, new[] { "qalam" }, 3, PredictionRankingPolicy.Current));
    }

    [Fact]
    public void CandidatePolicyLetsDictionaryBeatWeakContextOnly()
    {
        var weakPhrase = Candidate("qaror", 10);
        var strongPhrase = Candidate("qaror", 80);
        Assert.Equal(new[] { "qalam", "qaror" }, SuggestionOrdering.Merge("qa",
            new[] { weakPhrase }, new[] { "qalam" }, 3, PredictionRankingPolicy.AccuracyCandidate));
        Assert.Equal(new[] { "qaror", "qalam" }, SuggestionOrdering.Merge("qa",
            new[] { strongPhrase }, new[] { "qalam" }, 3, PredictionRankingPolicy.AccuracyCandidate));
    }

    [Fact]
    public async Task AfterSpaceGateCanBeCalibratedWithoutChangingCurrentPolicy()
    {
        var model = new PredictionCollection { Id = "legal", Sources = new() {
            new PredictionSource { Path = "source", Sequences = new() {
                ["sud qaror"] = 4, ["sud ajrim"] = 3
            } }
        } };
        var request = new PredictionRequest("sud ", "", ScriptType.Latin, new[] { "legal" }, 1);
        Assert.NotEmpty(await new PhrasePredictionEngine(new[] { model }).PredictAsync(request));
        var strict = new PhrasePredictionEngine(new[] { model }, 0.4, 0, true,
            PredictionRankingPolicy.AccuracyCandidate, 0.6, 0.1);
        Assert.Empty(await strict.PredictAsync(request));
    }

    private static PredictionCandidate Candidate(string text, double score) =>
        new PredictionCandidate(text.Substring(2), text, score, 2, 1, 1,
            Array.Empty<PredictionProvenance>());
}
