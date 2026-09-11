using UzbekOrfoAddIn.Services;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class TransliterationServiceTests
{
    [Theory]
    [InlineData("kitob", "китоб")]
    [InlineData("o'zbek", "ўзбек")]
    [InlineData("o\u02BBzbek", "ўзбек")]
    [InlineData("g'isht", "ғишт")]
    public void ToCyrillic_ConvertsCommonUzbekWords(string latin, string expectedCyrillic)
    {
        using var fixture = CreateService();
        Assert.Equal(expectedCyrillic, fixture.Service.ToCyrillic(latin));
    }

    [Theory]
    [InlineData("китоб", "kitob")]
    [InlineData("ўзбек", "o\u02BBzbek")]
    [InlineData("ғишт", "g\u02BBisht")]
    public void ToLatin_ConvertsCommonUzbekWords(string cyrillic, string expectedLatin)
    {
        using var fixture = CreateService();
        Assert.Equal(expectedLatin, fixture.Service.ToLatin(cyrillic));
    }

    private static TemporaryTransliterationService CreateService()
    {
        string exceptionsPath = Path.Combine(
            Path.GetTempPath(),
            $"UzbekOrfo-test-{Guid.NewGuid():N}.json");

        return new TemporaryTransliterationService(exceptionsPath);
    }

    private sealed class TemporaryTransliterationService : IDisposable
    {
        private readonly string _exceptionsPath;

        public TransliterationService Service { get; }

        public TemporaryTransliterationService(string exceptionsPath)
        {
            _exceptionsPath = exceptionsPath;
            Service = new TransliterationService(exceptionsPath);
        }

        public void Dispose()
        {
            if (File.Exists(_exceptionsPath))
                File.Delete(_exceptionsPath);
        }
    }
}
