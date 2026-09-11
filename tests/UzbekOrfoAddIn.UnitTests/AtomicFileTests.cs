using UzbekOrfoAddIn.Helpers;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class AtomicFileTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "UzbekOrfo-tests-" + Guid.NewGuid().ToString("N"));
    public void Dispose()
    {
        string fullPath = Path.GetFullPath(directory);
        string testPrefix = Path.Combine(Path.GetFullPath(Path.GetTempPath()), "UzbekOrfo-tests-");
        if (!fullPath.StartsWith(testPrefix, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Test cleanup escaped its temporary directory.");
        if (Directory.Exists(fullPath)) Directory.Delete(fullPath, true);
    }

    [Fact]
    public void SavingExistingFile_PreservesPreviousVersionAsBackup()
    {
        string path = Path.Combine(directory, "settings.cfg");
        AtomicFile.WriteAllLines(path, new[] { "old" });
        AtomicFile.WriteAllLines(path, new[] { "new" });
        Assert.Equal(new[] { "new" }, File.ReadAllLines(path));
        Assert.Equal(new[] { "old" }, File.ReadAllLines(path + ".bak"));
        Assert.Empty(Directory.GetFiles(directory, "*.tmp"));
    }

    [Fact]
    public async Task ConcurrentSaves_ProduceOneCompleteSnapshot()
    {
        string path = Path.Combine(directory, "settings.cfg");
        await Task.WhenAll(Enumerable.Range(0, 20).Select(i => Task.Run(() =>
            AtomicFile.WriteAllLines(path, Enumerable.Repeat(i.ToString(), 100).ToArray()))));
        var lines = File.ReadAllLines(path);
        Assert.Equal(100, lines.Length);
        Assert.Single(lines.Distinct());
        Assert.Empty(Directory.GetFiles(directory, "*.tmp"));
    }

    [Fact]
    public void FailedReplacement_LeavesOriginalFileIntact()
    {
        string path = Path.Combine(directory, "settings.cfg");
        AtomicFile.WriteAllLines(path, new[] { "original" });
        Directory.CreateDirectory(path + ".bak"); // An invalid backup destination forces replacement to fail.
        Assert.ThrowsAny<IOException>(() => AtomicFile.WriteAllLines(path, new[] { "new" }));
        Assert.Equal(new[] { "original" }, File.ReadAllLines(path));
        Assert.Empty(Directory.GetFiles(directory, "*.tmp"));
    }
}
