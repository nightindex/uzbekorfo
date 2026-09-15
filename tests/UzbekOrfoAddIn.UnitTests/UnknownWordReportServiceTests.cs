using UzbekOrfoAddIn.Services;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class UnknownWordReportServiceTests
{
    [Fact]
    public void RecordBatch_PersistsOnlyAggregatedWordsWithoutContext()
    {
        string directory = Path.Combine(Path.GetTempPath(), "UzbekOrfoTests", Guid.NewGuid().ToString("N"));
        string reportPath = Path.Combine(directory, "unknown_words.tsv");
        try
        {
            var service = new UnknownWordReportService(reportPath);
            service.RecordBatch(new[] { "Китоп", "китоп", "хатолик" });

            var reloaded = new UnknownWordReportService(reportPath).GetSnapshot();
            Assert.Equal(2, reloaded.Count);
            Assert.Equal(2, reloaded.Single(e => e.Word == "китоп").Count);
            string stored = File.ReadAllText(reportPath);
            Assert.DoesNotContain("context", stored, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("document", stored, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }
}
