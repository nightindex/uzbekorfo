using UzbekOrfoAddIn.Helpers;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class TextHelperTests
{
    [Theory]
    [InlineData("  Hello!  ", "hello")]
    [InlineData("...word...", "word")]
    [InlineData("   ", "")]
    public void NormalizeWord_TrimsAndRemovesEdgePunctuation(string input, string expected)
    {
        Assert.Equal(expected, TextHelper.NormalizeWord(input));
    }

    [Theory]
    [InlineData("o'zbek")]
    [InlineData("o\u02BBzbek")]
    [InlineData("o\u02BCzbek")]
    [InlineData("o\u2018zbek")]
    [InlineData("o\u2019zbek")]
    public void NormalizeWord_UsesOneKeyForUzbekApostropheVariants(string input)
    {
        Assert.Equal("o'zbek", TextHelper.NormalizeWord(input));
    }

    [Fact]
    public void Tokenize_KeepsCurlyApostropheWordsIntact()
    {
        var tokens = TextHelper.Tokenize("O\u2018zbek va g\u02BCisht");

        Assert.Equal(new[] { "o'zbek", "va", "g'isht" }, tokens.Select(token => token.Normalized));
    }

    [Fact]
    public void EditDistance_RecognizesTransposition()
    {
        Assert.Equal(1, TextHelper.EditDistance("kitob", "kiotb"));
    }

    [Fact]
    public void BoundedEditDistance_ReturnsThresholdPlusOneWhenOutsideBound()
    {
        Assert.Equal(2, TextHelper.BoundedEditDistance("abc", "uvwxyz", 1));
    }

    [Fact]
    public void UzbekStringComparer_UsesDigraphAwareLatinOrder()
    {
        Assert.True(UzbekStringComparer.Instance.Compare("ch", "d") < 0);
        Assert.True(UzbekStringComparer.Instance.Compare("n", "ng") < 0);
    }
}
