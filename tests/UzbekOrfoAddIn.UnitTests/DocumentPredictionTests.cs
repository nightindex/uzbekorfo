using System.IO.Compression;
using System.Text;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Prediction;
using UzbekOrfoAddIn.Services;
using Xunit;

namespace UzbekOrfoAddIn.UnitTests;

public sealed class DocumentPredictionTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), "matnai-tests-" + Guid.NewGuid().ToString("N"));
    public DocumentPredictionTests() { Directory.CreateDirectory(root); }
    public void Dispose() { Directory.Delete(root, true); }
    private string FileText(string name, string text) { var path = Path.Combine(root, name); File.WriteAllText(path, text, Encoding.UTF8); return path; }
    private static PredictionCollection Model(string id = "test")
    {
        var source = CollectionImporter.BuildSource("source.txt", new[] {
            "Суд томонидан қарор қабул қилинди бугун.", "Суд томонидан қарор қабул қилинди кеча.",
            "Суд томонидан қарор қабул қилинди яна.", "Sud tomonidan qaror qabul qilindi bugun.",
            "Sud tomonidan qaror qabul qilindi kecha.", "Oʻzbekiston qonunlari amal qiladi.", "Oʻzbekiston qonunlari yangilandi." });
        return new PredictionCollection { Id = id, Name = id, Sources = new() { source } };
    }

    [Fact] public async Task CompactIndexSharesOnlyIdenticalImmutableProvenanceAndDetachesFromInput()
    {
        var model = new PredictionCollection { Id = "compact", Name = "Original", Sources = new() {
            new PredictionSource { Path = "one.txt", Hash = "first", Sequences = new() {
                ["kitob"] = 3, ["kitoblar"] = 3, ["kitobcha"] = 2 } },
            new PredictionSource { Path = "two.txt", Hash = "second", Sequences = new() { ["kitob"] = 3 } }
        } };
        var engine = new PhrasePredictionEngine(new[] { model });
        var request = new PredictionRequest("", "ki", ScriptType.Latin, new[] { "compact" }, 42, 10);
        var before = await engine.PredictAsync(request);
        var book = before.Single(c => c.FullCompletion == "kitob");
        var books = before.Single(c => c.FullCompletion == "kitoblar");
        var booklet = before.Single(c => c.FullCompletion == "kitobcha");
        Assert.Equal(6, book.Support);
        Assert.Same(book.Provenance[0], books.Provenance[0]);
        Assert.NotSame(book.Provenance[0], book.Provenance[1]);
        Assert.NotSame(book.Provenance[0], booklet.Provenance[0]);
        Assert.Equal(2, booklet.Provenance[0].Support);
        model.Name = "Changed";
        model.Sources[0].Path = "changed.txt";
        model.Sources[0].Sequences.Clear(); model.Sources.Clear();
        var after = await engine.PredictAsync(request);
        Assert.Equal(before.Select(c => (c.FullCompletion, c.Score, c.Support)),
            after.Select(c => (c.FullCompletion, c.Score, c.Support)));
        Assert.All(after.SelectMany(c => c.Provenance), p => Assert.Equal("Original", p.CollectionName));
        Assert.Equal("one.txt", after.Single(c => c.FullCompletion == "kitob").Provenance[0].SourcePath);
    }
    [Fact] public void HtmlDocSkipsCommentsScriptsDeletedAndHeaderText()
    {
        var path = FileText("sample.doc", "<html><body><header>header</header><script>bad()</script><p>Ўзбекистон <b>қонуни</b></p><del>deleted</del><table><tr><td>Суд</td><td>қарори</td></tr></table></body></html>");
        Assert.Equal(new[] { "Ўзбекистон қонуни", "Суд", "қарори" }, new DocumentExtractor().Extract(path).Passages);
    }
    [Fact] public void DocxBodySkipsDeletedRevisionCommentsAndExternalReferences()
    {
        string path = Path.Combine(root, "body.docx");
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            using var writer = new StreamWriter(archive.CreateEntry("word/document.xml").Open());
            writer.Write("<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:p><w:r><w:t>Суд </w:t></w:r><w:del><w:r><w:delText>deleted</w:delText></w:r></w:del><w:r><w:t>қарори</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>мажлис</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>");
        }
        Assert.Equal(new[] { "Суд қарори", "мажлис" }, new DocumentExtractor().Extract(path).Passages);
    }
    [Fact] public void TextImportKeepsScriptsAndParagraphBoundaries()
    {
        Assert.Equal(new[] { "Суд қарори", "Sud qarori" }, new DocumentExtractor().Extract(FileText("both.txt", "Суд қарори\nSud qarori")).Passages);
    }
    [Fact] public async Task BuiltInLegalArtifactsLoadTogetherAndPredictInTheTypedScript()
    {
        var store = new CollectionStore(Path.Combine(root, "legal-artifacts"));
        foreach (string file in new[] { "legal_prediction.collection.gz", "legal_prediction_latin.collection.gz" })
        {
            using var stream = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Data", file));
            using var gzip = new GZipStream(stream, CompressionMode.Decompress);
            using var reader = new StreamReader(gzip, Encoding.UTF8);
            var envelope = Newtonsoft.Json.Linq.JObject.Parse(reader.ReadToEnd());
            Assert.Equal(CollectionStore.FormatVersion, (int)envelope["Version"]!);
            var collection = envelope["Collection"]!.ToObject<PredictionCollection>()!;
            Assert.True(collection.BuiltIn);
            Assert.Equal(collection.Id == "legal-latin" ? 154 : 155, collection.Sources.Count);
            Assert.Equal(57, collection.Sources.Count(s => !s.Path.Contains("/sud-court/")));
            Assert.All(collection.Sources, source => Assert.False(string.IsNullOrWhiteSpace(source.SourceUrl)));
            Assert.Equal(collection.Sources.Count, collection.Sources.Select(s => s.Path).Distinct().Count());
            store.Save(collection); // Includes the production collection validation.
        }
        var models = store.LoadAll();
        Assert.Empty(store.LoadErrors);
        Assert.Equal(new[] { "legal-default", "legal-latin" }, models.Select(m => m.Id).OrderBy(id => id));
        var latin = models.Single(m => m.Id == "legal-latin");
        Assert.All(latin.Sources, source => {
            Assert.StartsWith("https://lex.uz/docs/-", source.SourceUrl);
            Assert.All(source.Sequences.Keys, sequence => Assert.False(sequence.Any(TextHelper.IsCyrillic)));
        });
        var engine = new PhrasePredictionEngine(models);
        var ids = models.Select(m => m.Id).ToArray();
        foreach (var example in new[] {
            (Context: "", Prefix: "qo", Id: "legal-latin", Script: ScriptType.Latin),
            (Context: "O‘zbekiston ", Prefix: "", Id: "legal-latin", Script: ScriptType.Latin),
            (Context: "O'zbekiston ", Prefix: "", Id: "legal-latin", Script: ScriptType.Latin),
            (Context: "", Prefix: "қо", Id: "legal-default", Script: ScriptType.Cyrillic),
            (Context: "Ўзбекистон ", Prefix: "", Id: "legal-default", Script: ScriptType.Cyrillic) })
        {
            var results = await engine.PredictAsync(new PredictionRequest(example.Context, example.Prefix, ScriptType.Unknown, ids, 17));
            Assert.NotEmpty(results);
            Assert.All(results, candidate => {
                Assert.All(candidate.Provenance, p => Assert.Equal(example.Id, p.CollectionId));
                Assert.InRange(candidate.FullCompletion.Split(' ').Length, 1, 5);
                Assert.Equal(example.Prefix + candidate.InsertableTail, candidate.FullCompletion);
                Assert.False(candidate.FullCompletion.Any(example.Script == ScriptType.Latin
                    ? TextHelper.IsCyrillic : TextHelper.IsLatin));
            });
        }
    }
    [Fact] public void GenuineWord97BinaryIsExtracted()
    {
        var extracted = new DocumentExtractor().Extract(Path.Combine(AppContext.BaseDirectory, "Corpora", "word97-simple.doc"));
        Assert.NotEmpty(extracted.Passages);
        Assert.Contains("This is a simple file", string.Join(" ", extracted.Passages));
    }
    [Fact] public void MalformedDocxFailsWithoutLaunchingWord()
    { Assert.Throws<InvalidDataException>(() => new DocumentExtractor().Extract(FileText("broken.docx", "not zip"))); }
    [Fact] public async Task DeduplicationAndMissingRefreshPreserveGoodSources()
    {
        var path = FileText("first.txt", "Sud tomonidan qaror. Sud tomonidan ajrim.");
        var duplicate = FileText("copy.txt", File.ReadAllText(path));
        var importer = new CollectionImporter();
        var first = await importer.ImportAsync(new PredictionCollection(), new[] { path, duplicate });
        Assert.Single(first.Files.Where(f => f.Status == CollectionImportStatus.Duplicate));
        Assert.Empty(first.Collection.Sources[1].Sequences);
        var counts = first.Collection.Sources[0].Sequences.ToArray();
        File.Delete(path);
        var refreshed = await importer.ImportAsync(first.Collection, new[] { path });
        Assert.Equal(CollectionImportStatus.Missing, refreshed.Files[0].Status);
        Assert.Equal(counts, refreshed.Collection.Sources[0].Sequences.ToArray());
        refreshed.Collection.Sources.RemoveAt(0); CollectionImporter.Recount(refreshed.Collection);
        Assert.NotEmpty(refreshed.Collection.Sources[0].Sequences); // duplicate now owns its contribution
    }
    [Fact] public async Task CancelledImportDoesNotMutateOriginal()
    {
        var original = Model(); using var cancel = new CancellationTokenSource(); cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new CollectionImporter().ImportAsync(original,
            new[] { FileText("new.txt", "new word") }, cancellation: cancel.Token));
        Assert.Single(original.Sources); Assert.Equal("source.txt", original.Sources[0].Path);
    }
    [Fact] public async Task NewParagraphDoesNotCreateAnArtificialPhrase()
    {
        var source = CollectionImporter.BuildSource("x", new[] { "Sud qarori.", "Qabul qilindi." });
        Assert.DoesNotContain("qarori qabul", source.Sequences.Keys);
        var result = await new PhrasePredictionEngine(Model().Yield()).PredictAsync(new PredictionRequest("unrelated. ", "", ScriptType.Latin, new[] { "test" }, 1));
        Assert.Empty(result);
    }
    [Theory] [InlineData("Суд томонидан ", "қа", ScriptType.Cyrillic, "қарор")]
    [InlineData("Sud tomonidan ", "qa", ScriptType.Latin, "qaror")]
    [InlineData("Суд томонидан ", "", ScriptType.Cyrillic, "қарор")]
    public async Task PredictsContextualWordsAndPhrases(string context, string prefix, ScriptType script, string expected)
    {
        var result = await new PhrasePredictionEngine(new[] { Model() }).PredictAsync(new PredictionRequest(context, prefix, script, new[] { "test" }, 42));
        Assert.NotEmpty(result); Assert.StartsWith(expected, result[0].FullCompletion);
        Assert.All(result, c => { Assert.InRange(c.FullCompletion.Split(' ').Length, 1, 5); Assert.Equal(42, c.RequestId); Assert.Equal(prefix + c.InsertableTail, c.FullCompletion); });
    }
    [Fact] public async Task DisabledCollectionsAndCrossScriptResultsAreExcluded()
    {
        var engine = new PhrasePredictionEngine(new[] { Model() });
        Assert.Empty(await engine.PredictAsync(new PredictionRequest("Sud tomonidan ", "", ScriptType.Latin, Array.Empty<string>(), 1)));
        Assert.Empty(await engine.PredictAsync(new PredictionRequest("Sud tomonidan ", "қа", ScriptType.Latin, new[] { "test" }, 2)));
    }
    [Fact] public async Task AmbiguousAfterSpaceSuggestionsStayQuietButTypedPrefixesStillComplete()
    {
        var model = new PredictionCollection { Id = "ambiguous", Sources = new() {
            new PredictionSource { Path = "fixture", Sequences = new() {
                ["sud qaror"] = 3, ["sud ajrim"] = 3, ["sud qaror qabul"] = 3
            } }
        } };
        var request = new PredictionRequest("Sud ", "", ScriptType.Latin, new[] { model.Id }, 1);
        request.AcceptanceCounts = new Dictionary<string, int> { [model.Id + "\tqaror"] = 20 };
        Assert.NotEmpty(await new PhrasePredictionEngine(new[] { model }, selectiveNextWord: false).PredictAsync(request));
        Assert.Empty(await new PhrasePredictionEngine(new[] { model }).PredictAsync(request));
        request.Prefix = "qa";
        Assert.NotEmpty(await new PhrasePredictionEngine(new[] { model }).PredictAsync(request));
        request.Prefix = "";
        model.Sources[0].Sequences["sud qaror"] = 9;
        var confident = await new PhrasePredictionEngine(new[] { model }).PredictAsync(request);
        Assert.NotEmpty(confident);
        Assert.All(confident, candidate => Assert.StartsWith("qaror", candidate.FullCompletion));
    }
    [Fact] public async Task DuplicateSourcesCannotInventRepeatedPhraseEvidence()
    {
        var source = CollectionImporter.BuildSource("x", new[] { "Sud tomonidan qaror qabul qilindi." });
        var model = new PredictionCollection { Id = "x", Sources = new() { source, source } };
        var result = await new PhrasePredictionEngine(new[] { model }).PredictAsync(new PredictionRequest("Sud tomonidan ", "", ScriptType.Latin, new[] { "x" }, 1));
        Assert.Empty(result);
    }
    [Fact] public async Task PrefixApostrophesAndUppercaseArePreserved()
    {
        var engine = new PhrasePredictionEngine(new[] { Model() });
        var lower = await engine.PredictAsync(new PredictionRequest("", "O‘", ScriptType.Latin, new[] { "test" }, 1));
        Assert.Contains(lower, c => c.FullCompletion.StartsWith("O‘"));
        var upper = await engine.PredictAsync(new PredictionRequest("Sud tomonidan ", "QA", ScriptType.Latin, new[] { "test" }, 2));
        Assert.All(upper, c => Assert.Equal(c.FullCompletion.ToUpperInvariant(), c.FullCompletion));
    }
    [Fact] public void PrivateModelsAreEncryptedRecoverFromBackupAndDeleteWithoutSources()
    {
        var store = new CollectionStore(Path.Combine(root, "models")); var model = Model();
        string source = FileText("keep.txt", "keep source"); model.Sources[0].Path = source;
        store.Save(model); model.Name = "updated"; store.Save(model);
        string path = Directory.GetFiles(store.Root, "*.dpapi").Single();
        Assert.DoesNotContain("Суд", Encoding.UTF8.GetString(File.ReadAllBytes(path)));
        File.WriteAllBytes(path, new byte[] { 0, 1, 2 });
        Assert.Single(store.LoadAll()); Assert.NotEmpty(store.LoadErrors);
        store.Delete(model.Id); Assert.Empty(store.LoadAll()); Assert.True(File.Exists(source));
        Assert.Empty(Directory.GetFiles(store.Root, "*.bak"));
    }
    [Fact] public void StoreRejectsTraversalAndUnknownVersions()
    {
        var store = new CollectionStore(root); Assert.Throws<ArgumentException>(() => store.Delete("../outside"));
        Assert.Throws<ArgumentException>(() => store.Save(new PredictionCollection { Id = "../bad" }));
    }
    [Fact] public void CancelledPublicationPreservesPreviousModelAndDeletionCleansStaging()
    {
        var store = new CollectionStore(Path.Combine(root, "models")); var model = Model();
        store.Save(model); string originalName = model.Name;
        using var cancel = new CancellationTokenSource(); cancel.Cancel(); model.Name = "cancelled";
        Assert.ThrowsAny<OperationCanceledException>(() => store.Save(model, cancellation: cancel.Token));
        Assert.Equal(originalName, store.LoadAll().Single().Name);
        string primary = Directory.GetFiles(store.Root, "*.dpapi").Single();
        string staging = primary + "." + Guid.NewGuid().ToString("N") + ".tmp";
        File.Copy(primary, staging);
        store.Delete(model.Id);
        Assert.False(File.Exists(staging)); Assert.Empty(store.LoadAll());
    }
    [Fact] public async Task AfterSpaceScriptUsesNearbyWordNotEarlierLatinAcronym()
    {
        var results = await new PhrasePredictionEngine(new[] { Model() }).PredictAsync(
            new PredictionRequest("ABC Суд томонидан ", "", ScriptType.Unknown, new[] { "test" }, 1));
        Assert.NotEmpty(results); Assert.StartsWith("қарор", results[0].FullCompletion);
    }
    [Fact] public async Task AcceptanceBoostCannotEnableAnInactiveCollection()
    {
        var request = new PredictionRequest("Sud tomonidan ", "qa", ScriptType.Latin, Array.Empty<string>(), 1);
        request.AcceptanceCounts = new Dictionary<string, int> { ["test\tqaror"] = int.MaxValue };
        Assert.Empty(await new PhrasePredictionEngine(new[] { Model() }).PredictAsync(request));
    }
    [Fact] public void LegacyAcceptanceMigrationIsVersionedAndResetDoesNotRestoreBackup()
    {
        string legacy = Path.Combine(root, "old.tsv"); var old = new CompletionPreferences(legacy);
        old.Record("salom"); old.Save();
        var store = new AcceptanceStore(root); store.Load(legacy);
        Assert.Equal(1, store.DictionarySnapshot()["salom"]);
        store.Record("private", "qaror qabul qilindi"); store.RemoveMissingCollections(Array.Empty<string>());
        Assert.DoesNotContain(store.Snapshot().Keys, k => k.StartsWith("private\t"));
        store.Clear(legacy); store.Load(legacy); Assert.Empty(store.Snapshot()); Assert.False(File.Exists(legacy + ".bak"));
    }
    [Fact] public void MetricsAreOptInAndContainNoText()
    {
        var metrics = new PredictionMetrics(root); metrics.Count("shown"); metrics.Save();
        Assert.False(File.Exists(Path.Combine(root, "metrics-v1.tsv")));
        metrics.Enabled = true; metrics.Count("shown"); metrics.Latency(45); metrics.Save();
        Assert.Contains("shown\t1", File.ReadAllText(Path.Combine(root, "metrics-v1.tsv")));
    }
    [Theory]
    [InlineData("Ўзбекистон ", "", "республикаси", "Республикаси")]
    [InlineData("Ўзбекистон ", "Р", "Республикасининг қонуни", "Республикасининг қонуни")]
    [InlineData("", "Ўз", "Ўзбекистон республикаси", "Ўзбекистон Республикаси")]
    [InlineData("Oʻzbekiston ", "", "respublikasi", "Respublikasi")]
    [InlineData("O‘zbekiston ", "", "respublikasining qonuni", "Respublikasining qonuni")]
    [InlineData("бошқа ", "", "республикаси", "республикаси")]
    [InlineData("Ўзбекистон. ", "", "республикаси", "республикаси")]
    [InlineData("Ўзбекистон ", "ре", "республикаси", "республикаси")]
    public void OfficialNameCasingRespectsContextAndAlreadyTypedText(string context, string prefix, string completion, string expected)
    { Assert.Equal(expected, PredictionCasing.Apply(context, prefix, completion)); }

    [Fact] public async Task CountryContextCapitalizesSuggestedRepublicWithoutUppercasingWholeTail()
    {
        var source = CollectionImporter.BuildSource("legal", new[] {
            "Ўзбекистон республикаси қонуни бугун.", "Ўзбекистон республикаси қонуни кеча." });
        var model = new PredictionCollection { Id = "legal", Sources = new() { source } };
        var engine = new PhrasePredictionEngine(new[] { model });
        foreach (string prefix in new[] { "", "Р", "Ре" })
        {
            var result = await engine.PredictAsync(new PredictionRequest("Ўзбекистон ", prefix, ScriptType.Cyrillic, new[] { "legal" }, 1));
            Assert.NotEmpty(result); Assert.StartsWith("Республикаси", result[0].FullCompletion);
            Assert.Equal(prefix + result[0].InsertableTail, result[0].FullCompletion);
        }
    }
}

internal static class PredictionTestEnumerable
{
    internal static IEnumerable<T> Yield<T>(this T value) { yield return value; }
}
