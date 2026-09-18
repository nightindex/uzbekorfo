using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Prediction;
using UzbekOrfoAddIn.Services;

// Developer tool: public corpus only. Never runs in Word and never imports personal collections.
internal static class Program
{
    private static bool latin;
    private static string CollectionId => latin ? "legal-latin" : "legal-default";
    private sealed class Document
    {
        public string Path, Url, Hash, Family, Split;
        public string[] Sentences;
        public PredictionSource Source;
    }
    private sealed class Evaluation
    {
        public int Cases, Offered, Top1, Top3, SavedCharacters, RemainingCharacters;
        public double Top1Rate => Cases == 0 ? 0 : (double)Top1 / Cases;
        public double Top3Rate => Cases == 0 ? 0 : (double)Top3 / Cases;
        public double Coverage => Cases == 0 ? 0 : (double)Offered / Cases;
        public double Top1WhenOffered => Offered == 0 ? 0 : (double)Top1 / Offered;
        public double SimulatedKeystrokesSaved => RemainingCharacters == 0 ? 0 : (double)SavedCharacters / RemainingCharacters;
    }
    private sealed class Case
    {
        public string Source, Context, Prefix, Expected;
        public string ExpectedTail;
    }
    private static async Task<int> Main(string[] args)
    {
        if (args.Length < 1 || args.Length > 4 || (args.Length >= 2 && args[1] != "latin" && args[1] != "cyrillic") ||
            args.Skip(2).Any(a => a != "--pilot-evaluation" && a != "--include-court"))
        { Console.Error.WriteLine("Usage: MatnAi.Corpus <repository-root> [latin|cyrillic] [--pilot-evaluation] [--include-court]"); return 2; }
        latin = args.Length >= 2 && args[1] == "latin";
        bool pilotEvaluation = args.Contains("--pilot-evaluation");
        bool includeCourt = args.Contains("--include-court");
        string root = Path.GetFullPath(args[0]);
        string raw = Path.Combine(root, "corpus", latin ? "raw-latin" : "raw_cyrillic");
        // Keep older local snapshots usable without renaming or copying sources.
        if (!latin && !Directory.Exists(raw)) raw = Path.Combine(root, "corpus", "raw");
        string output = Path.Combine(root, "src", "UzbekOrfoAddIn", "Data", latin ? "legal_prediction_latin.collection.gz" : "legal_prediction.collection.gz");
        string reports = Path.Combine(root, "docs", "prediction", latin ? "latin" : ""); Directory.CreateDirectory(reports);
        // Expanded models are candidates, never an implicit production replacement.
        if (includeCourt)
        {
            reports = Path.Combine(reports, "expanded"); Directory.CreateDirectory(reports);
            string candidates = Path.Combine(root, "corpus", "candidates"); Directory.CreateDirectory(candidates);
            output = Path.Combine(candidates, Path.GetFileName(output));
        }
        var watch = Stopwatch.StartNew();
        var extractor = new DocumentExtractor(); var documents = new List<Document>(); var failures = new List<object>();
        var exclusions = new List<object>();
        var courtSources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (includeCourt)
        {
            string courtRoot = Path.Combine(raw, "sud-court");
            var manifest = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(Path.Combine(courtRoot, "manifest.json")));
            foreach (var entry in manifest["documents"] ?? throw new InvalidDataException("Missing court source manifest."))
            {
                if ((string)entry["status"] == "unavailable") continue;
                string relativePath = (string)entry["path"], url = (string)entry["url"];
                string expectedUrl = latin ? @"\Ahttps://lex\.uz/docs/-\d+\?type=doc\z" : @"\Ahttps://lex\.uz/docs/\d+\?type=doc\z";
                if (string.IsNullOrEmpty(relativePath) || !Regex.IsMatch(url ?? "", expectedUrl))
                    throw new InvalidDataException("Invalid court source provenance.");
                string sourcePath = Path.GetFullPath(Path.Combine(courtRoot, relativePath.Replace('\\', '/')));
                if (!sourcePath.StartsWith(courtRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) || !File.Exists(sourcePath))
                    throw new InvalidDataException("Court manifest source missing or outside corpus: " + relativePath);
                courtSources.Add(sourcePath, url);
            }
        }
        var sourceIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string manifest in Directory.GetFiles(Path.Combine(root, "docs"), "legal*corpus.md"))
        foreach (Match match in Regex.Matches(File.ReadAllText(manifest), @"`([^`]+\.doc)`\s*\|\s*(\d+)"))
            sourceIds[match.Groups[1].Value] = match.Groups[2].Value;
        var latinSources = new Dictionary<string, string>(StringComparer.Ordinal);
        if (latin)
        {
            var manifest = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(Path.Combine(raw, "manifest.json")));
            foreach (var entry in manifest["documents"] ?? throw new InvalidDataException("Missing Latin source manifest."))
            {
                string path = (string)entry["path"], url = (string)entry["url"];
                if (string.IsNullOrEmpty(path) || !Regex.IsMatch(url ?? "", @"\Ahttps://lex\.uz/docs/-\d+\?type=doc\z"))
                    throw new InvalidDataException("Invalid Latin source provenance.");
                latinSources.Add(path, url);
            }
        }
        // The separately downloaded court-materials snapshot is not part of the
        // frozen 57-source model until it passes provenance/review gates.
        var files = Directory.GetFiles(raw, "*.doc", SearchOption.AllDirectories)
            .Where(p => includeCourt || !p.Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries)
                .Contains("sud-court", StringComparer.OrdinalIgnoreCase))
            .OrderBy(p => p, StringComparer.Ordinal).ToArray();
        if (files.Length == 0 || (latin && files.Length != latinSources.Count + courtSources.Count))
            throw new InvalidDataException("Corpus files are empty or do not match the Latin manifest.");
        for (int index = 0; index < files.Length; index++)
        {
            string path = files[index]; string relative = Path.GetRelativePath(root, path).Replace('\\', '/');
            bool court = relative.Contains("/sud-court/");
            Console.WriteLine($"Extract {index + 1}/{files.Length}: {Path.GetFileName(path)}");
            try
            {
                var extracted = extractor.Extract(path);
                if (court && !courtSources.ContainsKey(path)) throw new InvalidDataException("Court file has no source manifest entry.");
                if (latin && (!extracted.Passages.Any(p => p.Any(TextHelper.IsLatin)) || extracted.Passages.Any(p => p.Any(TextHelper.IsCyrillic))))
                    throw new InvalidDataException("Expected Latin source text; found empty or Cyrillic content.");
                if (court && !latin && !extracted.Passages.Any(p => p.IndexOfAny("ўқғҳЎҚҒҲ".ToCharArray()) >= 0))
                    throw new InvalidDataException("No distinctive Uzbek Cyrillic letters; language review required.");
                var sentences = extracted.Passages.SelectMany(p => Regex.Split(p, @"[.!?;:\r\n\t\a\d()\[\]«»]+"))
                    .Select(s => string.Join(" ", TextHelper.Tokenize(s).Where(t => WordCompletionEngine.IsToken(t.Normalized)).Select(t => t.Normalized)))
                    .Where(s => s.Length > 0).Distinct(StringComparer.Ordinal).ToArray();
                string id; sourceIds.TryGetValue(Path.GetFileName(path), out id);
                if (id == null) id = Regex.Match(Path.GetFileName(path), @"lexuz-(\d+)").Groups[1].Value;
                var source = CollectionImporter.BuildSource(relative, extracted.Passages, 40000);
                source.SourceUrl = string.IsNullOrEmpty(id) ? null : "https://lex.uz/docs/" + id + "?type=doc";
                if (court) source.SourceUrl = courtSources[path];
                else if (latin) source.SourceUrl = latinSources[Path.GetRelativePath(raw, path).Replace('\\', '/')];
                var document = new Document { Path = relative, Url = source.SourceUrl, Hash = source.Hash, Source = source, Sentences = sentences };
                document.Family = id ?? source.Hash;
                if (Path.GetFileName(path).Contains("fuqarolik-kodeksi")) document.Family = "civil-code-parts";
                if (string.IsNullOrEmpty(document.Family)) document.Family = source.Hash;
                documents.Add(document);
            }
            catch (Exception ex)
            {
                var failure = new { source = relative, error = ex.Message };
                if (court) exclusions.Add(failure); else failures.Add(failure);
            }
        }
        if (failures.Count > 0)
        {
            File.WriteAllText(Path.Combine(reports, "extraction-failures.json"), JsonConvert.SerializeObject(failures, Formatting.Indented));
            Console.Error.WriteLine("Extraction failed; previous model retained. See extraction-failures.json.");
            return 1;
        }
        // Group exact duplicates and near-identical versions BEFORE splitting.
        var groups = Enumerable.Range(0, documents.Count).ToArray();
        int Find(int i) { while (groups[i] != i) i = groups[i]; return i; }
        for (int i = 0; i < documents.Count; i++)
        for (int j = 0; j < i; j++)
        {
            var a = new HashSet<string>(documents[i].Sentences); var b = new HashSet<string>(documents[j].Sentences);
            int shared = a.Count(s => b.Contains(s));
            if (documents[i].Hash == documents[j].Hash || documents[i].Family == documents[j].Family ||
                shared >= 10 && shared / (double)Math.Max(1, Math.Min(a.Count, b.Count)) >= 0.5) groups[Find(i)] = Find(j);
        }
        var ordered = documents.Select((d, i) => new { d, group = Find(i) }).GroupBy(x => x.group)
            .OrderBy(g => CollectionImporter.Hash(string.Join("|", g.Select(x => x.d.Hash).OrderBy(h => h)))).ToArray();
        int holdout = Math.Max(1, ordered.Length / 10);
        for (int i = 0; i < ordered.Length; i++)
            foreach (var row in ordered[i]) { row.d.Family = "group-" + i; row.d.Split = i < holdout ? "test" : i < holdout * 2 ? "validation" : "train"; }
        var train = MakeModel(documents.Where(d => d.Split == "train"));
        var trainingSentences = new HashSet<string>(documents.Where(d => d.Split == "train").SelectMany(d => d.Sentences));
        var validation = Cases(documents.Where(d => d.Split == "validation"), trainingSentences);
        var test = Cases(documents.Where(d => d.Split == "test"), trainingSentences);
        var dictionaryPath = Path.Combine(root, "src", "UzbekOrfoAddIn", "Data", "uzbek_main.dic");
        var lexical = new WordCompletionEngine(File.ReadLines(dictionaryPath));
        if (pilotEvaluation)
        {
            var legacy = new PhrasePredictionEngine(new[] { train }, selectiveNextWord: false);
            var selective = new PhrasePredictionEngine(new[] { train });
            async Task<object> Compare(List<Case> samples) => new {
                legacy = await Evaluate(legacy, samples, lexical, false),
                selective = await Evaluate(selective, samples, lexical, false)
            };
            var pilot = new {
                schema = "matnai-pilot-evaluation-v1", generatedUtc = DateTime.UtcNow,
                script = latin ? "Latin" : "Cyrillic",
                policy = "After-space next-word support >=2, probability >=0.4, lead over runner-up >=0.1. Applies before acceptance boosts; prefix behavior unchanged.",
                validationAfterSpace = await Compare(validation.Where(c => c.Prefix.Length == 0).ToList()),
                testAfterSpace = await Compare(test.Where(c => c.Prefix.Length == 0).ToList()),
                testPrefix = await Compare(test.Where(c => c.Prefix.Length > 0).ToList()),
                combinedModels = await BenchmarkCombinedModels(root),
                notes = "Same grouped split and normalized replay for both policies. Exact match is not human usefulness. Training-only engine for replay; all-source packaged engines for performance. No review records or model files overwritten.",
                independentReviewPassed = false
            };
            string pilotPath = Path.Combine(reports, "pilot-report.json");
            File.WriteAllText(pilotPath, JsonConvert.SerializeObject(pilot, Formatting.Indented), new UTF8Encoding(false));
            Console.WriteLine("Pilot evaluation: " + pilotPath);
            return 0;
        }
        var calibration = new List<object>(); double selectedProbability = 0.25, selectedMargin = 0.05, best = -1;
        foreach (double probability in new[] { 0.15, 0.25, 0.4 })
        foreach (double margin in new[] { 0.0, 0.05, 0.1 })
        {
            var engine = new PhrasePredictionEngine(new[] { train }, probability, margin);
            var evaluation = await Evaluate(engine, validation, lexical, false);
            calibration.Add(new { probability, margin, evaluation });
            if (evaluation.SimulatedKeystrokesSaved > best)
            { best = evaluation.SimulatedKeystrokesSaved; selectedProbability = probability; selectedMargin = margin; }
        }
        // Evaluate the same frozen thresholds deployed by Word. Calibration remains
        // advisory until a separately reviewed change updates runtime defaults.
        var evaluationEngine = new PhrasePredictionEngine(new[] { train });
        var contextual = await Evaluate(evaluationEngine, test, lexical, false);
        var baseline = await Evaluate(evaluationEngine, test, lexical, true);
        // Controlled comparison: both engines use the expanded split; no held-out
        // or related-version group is admitted to either training model.
        Evaluation originalCorpus = null;
        object qualitySlices = null;
        if (includeCourt)
        {
            var originalTrain = MakeModel(documents.Where(d => d.Split == "train" && !d.Path.Contains("/sud-court/")));
            var originalEngine = new PhrasePredictionEngine(new[] { originalTrain });
            originalCorpus = await Evaluate(originalEngine, test, lexical, false);
            async Task<object> CompareSlice(IEnumerable<Case> cases)
            {
                var samples = cases.ToList();
                return new { original = await Evaluate(originalEngine, samples, lexical, false),
                    expanded = await Evaluate(evaluationEngine, samples, lexical, false) };
            }
            qualitySlices = new {
                originalDocuments = await CompareSlice(test.Where(c => !c.Source.Contains("/sud-court/"))),
                courtDocuments = await CompareSlice(test.Where(c => c.Source.Contains("/sud-court/"))),
                afterSpace = await CompareSlice(test.Where(c => c.Prefix.Length == 0)),
                twoLetterPrefix = await CompareSlice(test.Where(c => c.Prefix.Length > 0))
            };
        }
        // Artifact uses all public sources AFTER unbiased held-out evaluation.
        var model = MakeModel(documents);
        CollectionStore.Validate(model);
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); long before = GC.GetTotalMemory(true);
        var runtimeEngine = new PhrasePredictionEngine(new[] { model });
        long engineBytes = Math.Max(0, GC.GetTotalMemory(true) - before);
        var latencies = new List<double>();
        foreach (var sample in test.Take(100)) await runtimeEngine.PredictAsync(Request(sample));
        foreach (var sample in test.Take(1000))
        { var timer = Stopwatch.StartNew(); await runtimeEngine.PredictAsync(Request(sample)); latencies.Add(timer.Elapsed.TotalMilliseconds); }
        latencies.Sort(); double p95 = latencies.Count == 0 ? 0 : latencies[(int)Math.Floor((latencies.Count - 1) * .95)];
        Console.WriteLine("Benchmark prediction while importing documents...");
        var concurrentImport = new CollectionImporter().ImportAsync(new PredictionCollection { Name = "Benchmark" }, files);
        var concurrentLatencies = new List<double>();
        foreach (var sample in test.Take(1000))
        { var timer = Stopwatch.StartNew(); await runtimeEngine.PredictAsync(Request(sample)); concurrentLatencies.Add(timer.Elapsed.TotalMilliseconds); }
        await concurrentImport;
        concurrentLatencies.Sort();
        double importP95 = concurrentLatencies.Count == 0 ? 0 : concurrentLatencies[(int)Math.Floor((concurrentLatencies.Count - 1) * .95)];
        // Publish only a completely extracted, validated and evaluated artifact.
        string staging = output + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var file = File.Create(staging))
            using (var gzip = new GZipStream(file, CompressionLevel.Optimal))
            using (var writer = new StreamWriter(gzip, new UTF8Encoding(false)))
                writer.Write(JsonConvert.SerializeObject(new { Version = CollectionStore.FormatVersion, Collection = model }));
            if (File.Exists(output)) File.Replace(staging, output, null);
            else File.Move(staging, output);
        }
        finally { if (File.Exists(staging)) File.Delete(staging); }
        var report = new {
            schema = "matnai-corpus-evaluation-v1", generatedUtc = DateTime.UtcNow, filesFound = files.Length,
            script = latin ? "Latin" : "Cyrillic", collectionId = CollectionId,
            filesExtracted = documents.Count, failures, exclusions, includeCourt, rawBytes = files.Sum(p => new FileInfo(p).Length),
            originalCorpusSameHoldout = originalCorpus,
            qualitySlices,
            comparisonScope = "Original-only and expanded training on the same expanded grouped split; not a comparison with historical report splits. Script checks are conservative heuristics, not linguistic approval.",
            extractedUniqueSentenceWords = documents.Sum(d => d.Source.WordCount),
            duplicateDocuments = documents.Count - documents.Select(d => d.Hash).Distinct().Count(),
            duplicateSentences = documents.Sum(d => d.Sentences.Length) - documents.SelectMany(d => d.Sentences).Distinct().Count(),
            modelRecords = model.Sources.Sum(s => s.Sequences.Count), artifactBytes = new FileInfo(output).Length,
            files = documents.Select(d => new { source = d.Path, url = d.Url, hash = d.Hash, family = d.Family, split = d.Split,
                words = d.Source.WordCount, retainedSequences = model.Sources.First(s => s.Path == d.Path).Sequences.Count }),
            validationCases = validation.Count, testCases = test.Count, calibration,
            selectedThresholds = new { probability = selectedProbability, margin = selectedMargin },
            evaluatedRuntimeThresholds = new { probability = 0.4, margin = 0.0 },
            metricsDefinition = "Deterministic held-out replay, alternating after-space and two-letter prefixes. Top-k measures exact next word; saved characters subtract one acceptance key and require the complete offered continuation to match. Shared training sentences are excluded. No independent linguistic labels.",
            contextual, dictionaryBaseline = baseline, warmQueryP95Ms = p95, duringImportQueryP95Ms = importP95, engineManagedBytes = engineBytes,
            processorCount = Environment.ProcessorCount, machineMemoryBytes = GC.GetGCMemoryInfo().TotalAvailableMemoryBytes,
            elapsedSeconds = watch.Elapsed.TotalSeconds, independentReviewPassed = false, officeMatrixPassed = false,
            releaseStatus = "experimental-awaiting-independent-review-and-office-matrix"
        };
        File.WriteAllText(Path.Combine(reports, "corpus-report.json"), JsonConvert.SerializeObject(report, Formatting.Indented), new UTF8Encoding(false));
        var review = test.GroupBy(c => c.Context + "|" + c.Prefix).Select(g => g.First())
            .OrderBy(c => CollectionImporter.Hash(c.Source + "|" + c.Context + "|" + c.Prefix)).Take(200)
            .Select((c, i) => new { id = i + 1, c.Source, c.Context, c.Prefix, c.Expected, reviewer = "", decision = "pending" });
        File.WriteAllText(Path.Combine(reports, "review-contexts.json"), JsonConvert.SerializeObject(review, Formatting.Indented), new UTF8Encoding(false));
        Console.WriteLine($"Extracted {documents.Count}/{files.Length}; model {report.modelRecords} sequences; test {test.Count}; top1 {contextual.Top1Rate:P1}; p95 {p95:F2}ms; managed index {engineBytes / 1048576.0:F1}MiB.");
        return failures.Count == 0 ? 0 : 1;
    }
    private static async Task<object> BenchmarkCombinedModels(string root)
    {
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        long before = GC.GetTotalMemory(true);
        var collections = new List<PredictionCollection>();
        foreach (string name in new[] { "legal_prediction.collection.gz", "legal_prediction_latin.collection.gz" })
        {
            using var stream = File.OpenRead(Path.Combine(root, "src", "UzbekOrfoAddIn", "Data", name));
            using var gzip = new GZipStream(stream, CompressionMode.Decompress);
            using var reader = new StreamReader(gzip, Encoding.UTF8);
            collections.Add(Newtonsoft.Json.Linq.JObject.Parse(reader.ReadToEnd())["Collection"].ToObject<PredictionCollection>());
        }
        var engine = new PhrasePredictionEngine(collections);
        long retainedBytes = Math.Max(0, GC.GetTotalMemory(true) - before);
        var ids = collections.Select(c => c.Id).ToArray();
        var requests = new[] {
            new PredictionRequest("Ўзбекистон ", "", ScriptType.Unknown, ids, 1),
            new PredictionRequest("O'zbekiston ", "", ScriptType.Unknown, ids, 2),
            new PredictionRequest("", "қа", ScriptType.Unknown, ids, 3),
            new PredictionRequest("", "qa", ScriptType.Unknown, ids, 4)
        };
        foreach (var request in requests) await engine.PredictAsync(request);
        var times = new List<double>();
        for (int i = 0; i < 1000; i++)
        {
            var timer = Stopwatch.StartNew();
            await engine.PredictAsync(requests[i % requests.Length]);
            times.Add(timer.Elapsed.TotalMilliseconds);
        }
        times.Sort();
        // The controller retains the index, not the deserialized source models.
        // Settings/import can retain source models, so report both states.
        collections.Clear();
        long indexOnlyBytes = Math.Max(0, GC.GetTotalMemory(true) - before);
        GC.KeepAlive(engine);
        return new { retainedModelsAndIndexBytes = retainedBytes, retainedIndexOnlyBytes = indexOnlyBytes, warmQueryP95Ms = times[949],
            processorCount = Environment.ProcessorCount, machineMemoryBytes = GC.GetGCMemoryInfo().TotalAvailableMemoryBytes,
            scope = ".NET 8 managed retained source models plus both phrase indexes, four query smoke workload. Excludes Word, dictionary, private collections, import and peak allocation. Not the reference PC or end-to-end typing latency." };
    }
    private static PredictionCollection MakeModel(IEnumerable<Document> documents)
    {
        // Detached model: do not mutate held-out sources, their passages, or extraction reports.
        var sources = documents.Select(d => new PredictionSource { Path = d.Path, SourceUrl = d.Url, Hash = d.Hash,
            WordCount = d.Source.WordCount, DisplayWords = new Dictionary<string, string>(d.Source.DisplayWords),
            Passages = d.Source.Passages, Sequences = new Dictionary<string, int>(d.Source.Sequences) }).ToList();
        var model = new PredictionCollection { Id = CollectionId, Name = latin ? "Ҳуқуқий ҳужжатлар (лотин)" : "Ҳуқуқий ҳужжатлар", BuiltIn = true, Sources = sources };
        CollectionImporter.Recount(model);
        int budget = 180000 / Math.Max(1, sources.Count);
        foreach (var source in sources)
        {
            source.Sequences = source.Sequences.OrderByDescending(p => p.Value).ThenBy(p => p.Key, StringComparer.Ordinal).Take(budget)
                .ToDictionary(p => p.Key, p => p.Value, StringComparer.Ordinal);
            source.Passages = new List<PredictionPassage>();
            source.DisplayWords = source.DisplayWords.Where(p => source.Sequences.ContainsKey(p.Key)).Take(200)
                .ToDictionary(p => p.Key, p => p.Value, StringComparer.Ordinal);
        }
        return model;
    }
    private static List<Case> Cases(IEnumerable<Document> documents, HashSet<string> trainingSentences)
    {
        var result = new List<Case>();
        foreach (var doc in documents)
        {
            var possible = new List<Case>();
            foreach (var sentence in doc.Sentences.Where(s => !trainingSentences.Contains(s)))
            {
                var words = sentence.Split(' ');
                for (int i = 2; i < words.Length; i += 3)
                {
                    if (words[i].Length < 3) continue;
                    string prefix = possible.Count % 2 == 0 ? "" : words[i].Substring(0, 2);
                    possible.Add(new Case { Source = doc.Path, Context = string.Join(" ", words.Skip(Math.Max(0, i - 6)).Take(Math.Min(6, i))) + " ",
                        Prefix = prefix, Expected = words[i], ExpectedTail = string.Join(" ", words.Skip(i).Take(5)).Substring(prefix.Length) });
                }
            }
            int stride = Math.Max(1, possible.Count / 300);
            result.AddRange(possible.Where((c, i) => i % stride == 0).Take(300));
        }
        return result;
    }
    private static PredictionRequest Request(Case sample) => new PredictionRequest(sample.Context, sample.Prefix,
        latin ? ScriptType.Latin : ScriptType.Cyrillic, new[] { CollectionId }, 1, 3);
    private static async Task<Evaluation> Evaluate(PhrasePredictionEngine engine, List<Case> cases, WordCompletionEngine lexical, bool baseline)
    {
        var result = new Evaluation();
        foreach (var sample in cases)
        {
            var words = baseline ? lexical.Complete(sample.Prefix, 3) :
                (await engine.PredictAsync(Request(sample))).Select(p => p.FullCompletion).Concat(lexical.Complete(sample.Prefix, 3)).Distinct().Take(3).ToArray();
            // Comparison uses the same apostrophe/case normalization as matching;
            // insertions still preserve display spelling in the actual engine.
            words = words.Select(w => string.Join(" ", TextHelper.Tokenize(w).Select(t => t.Normalized))).ToArray();
            result.Cases++; result.RemainingCharacters += sample.Expected.Length - sample.Prefix.Length;
            if (words.Length == 0) continue;
            result.Offered++;
            if (words[0].Split(' ')[0] == sample.Expected) result.Top1++;
            if (words.Any(w => w.Split(' ')[0] == sample.Expected)) result.Top3++;
            string tail = words[0].Substring(sample.Prefix.Length);
            if (sample.ExpectedTail == tail || sample.ExpectedTail.StartsWith(tail + " ", StringComparison.Ordinal))
                result.SavedCharacters += Math.Max(0, Math.Min(tail.Length, sample.Expected.Length - sample.Prefix.Length) - 1);
        }
        return result;
    }
}
