using System.Text;
using UzbekOrfoAddIn.Helpers;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class ImportTextReaderTests
{
    [Fact]
    public void ReadTextFile_ReadsUtf8WithoutBom()
    {
        string path = Path.Combine(Path.GetTempPath(), $"UzbekOrfo-import-{Guid.NewGuid():N}.txt");
        const string expected = "ўзбек\no'zbek";

        try
        {
            File.WriteAllBytes(path, new UTF8Encoding(false).GetBytes(expected));

            Assert.Equal(expected, ImportTextReader.ReadTextFile(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ParseDelimitedRecords_PreservesQuotedNewlinesAndDelimiters()
    {
        const string csv = "Word;Definition\nkitob;\"Bilim, manbai;\nIkkinchi satr\"";

        var records = ImportTextReader.ParseDelimitedRecords(csv);

        Assert.Equal(2, records.Count);
        Assert.Equal(new[] { "Word", "Definition" }, records[0]);
        Assert.Equal(new[] { "kitob", "Bilim, manbai;\nIkkinchi satr" }, records[1]);
    }
}
