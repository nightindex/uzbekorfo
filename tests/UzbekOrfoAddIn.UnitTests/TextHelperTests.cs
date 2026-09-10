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
