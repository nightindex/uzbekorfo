using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Services;

namespace UzbekOrfoAddIn.Prediction
{
    public sealed class CollectionImporter
    {
        public const int MaxFiles = 1000;
        public const long MaxJobBytes = 500L * 1024 * 1024;
        private static readonly SemaphoreSlim JobGate = new SemaphoreSlim(1, 1);
        private readonly DocumentExtractor _extractor = new DocumentExtractor();
        public static bool Supported(string path) => new[] { ".txt", ".docx", ".doc" }.Contains(Path.GetExtension(path).ToLowerInvariant());

        public static string[] Discover(IEnumerable<string> selections, bool recursive, CancellationToken cancellation = default(CancellationToken))
        {
            var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var pending = new Stack<string>();
            int directories = 0;
            foreach (string selection in selections)
            {
                string path = Path.GetFullPath(selection);
                if (Directory.Exists(path)) pending.Push(path);
                else files.Add(path);
            }
            while (pending.Count > 0)
            {
                cancellation.ThrowIfCancellationRequested();
                if (++directories > 10000) throw new InvalidDataException("Бир вазифада 10000 тагача папка кўриб чиқилади.");
                string directory = pending.Pop();
                if ((File.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0) continue;
                foreach (string file in Directory.EnumerateFiles(directory))
                {
                    if (Path.GetFileName(file).StartsWith("~$", StringComparison.Ordinal)) continue;
                    if (Supported(file)) files.Add(file);
                    if (files.Count > MaxFiles) throw new InvalidDataException("Бир вазифада 1000 тагача файл танланг.");
                }
                if (recursive)
                    foreach (string child in Directory.EnumerateDirectories(directory)) pending.Push(child);
            }
            if (files.Count > MaxFiles) throw new InvalidDataException("Бир вазифада 1000 тагача файл танланг.");
            return files.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();
        }

        /// <summary>Builds a new snapshot. Caller publishes only after success. Input is never mutated.</summary>
        public async Task<CollectionImportResult> ImportAsync(PredictionCollection original, IEnumerable<string> paths,
            IProgress<CollectionImportProgress> progress = null, CancellationToken cancellation = default(CancellationToken))
        {
            var selected = paths.Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (selected.Length > MaxFiles) throw new InvalidDataException("Бир вазифада 1000 тагача файл танланг.");
            await JobGate.WaitAsync(cancellation).ConfigureAwait(false);
            try { return await Task.Run(() => Import(original, selected, progress, cancellation), cancellation).ConfigureAwait(false); }
            finally { JobGate.Release(); }
        }

        private CollectionImportResult Import(PredictionCollection original, string[] paths,
            IProgress<CollectionImportProgress> progress, CancellationToken token)
        {
            var next = new PredictionCollection { Id = original.Id, Name = original.Name, BuiltIn = original.BuiltIn,
                Sources = original.Sources.Select(Clone).ToList(), UpdatedUtc = DateTime.UtcNow,
                ImportLocations = new Dictionary<string, bool>(original.ImportLocations, StringComparer.OrdinalIgnoreCase),
                ExcludedSources = new List<string>(original.ExcludedSources) };
            var report = new CollectionImportResult { Collection = next };
            int expectedSources = next.Sources.Select(s => s.Path).Concat(paths).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            if (expectedSources > MaxFiles) throw new InvalidDataException("Тўпламда 1000 тагача файл бўлиши мумкин.");
            int sourceBudget = Math.Max(100, (CollectionStore.MaxModelRecords - 1000) / Math.Max(1, expectedSources));
            long bytes = 0;
            for (int index = 0; index < paths.Length; index++)
            {
                token.ThrowIfCancellationRequested();
                string path = paths[index];
                var item = new CollectionImportFileReport { Path = path };
                try
                {
                    if (!Supported(path)) throw new NotSupportedException("Фақат DOCX, DOC ва TXT файллари қўлланади.");
                    var info = new FileInfo(path);
                    if (!info.Exists) { item.Status = CollectionImportStatus.Missing; item.Message = "Файл топилмади. Аввалги маълумот сақланди."; }
                    else if (info.Length > DocumentExtractor.MaxFileBytes || info.Length > MaxJobBytes - bytes)
                    { item.Status = CollectionImportStatus.LimitExceeded; item.Message = "50 МБ файл ёки 500 МБ вазифа чегараси ошди. Аввалги маълумот сақланди."; }
                    else
                    {
                        bytes += info.Length;
                        var source = BuildSource(path, _extractor.Extract(path, token).Passages, sourceBudget, token);
                        if (source.WordCount == 0) throw new InvalidDataException("Ўрганиш учун сўзлар топилмади.");
                        int oldIndex = next.Sources.FindIndex(s => string.Equals(s.Path, path, StringComparison.OrdinalIgnoreCase));
                        if (oldIndex >= 0) { source.SourceUrl = next.Sources[oldIndex].SourceUrl; next.Sources[oldIndex] = source; }
                        else next.Sources.Add(source);
                        item.WordCount = source.WordCount;
                        item.Status = next.Sources.Any(s => s != source && s.Hash == source.Hash) ? CollectionImportStatus.Duplicate : CollectionImportStatus.Imported;
                        item.Message = "Тайёр. Кам учрайдиган иборалар индекс ҳажмига мос равишда чекланади.";
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (NotSupportedException ex) { item.Status = CollectionImportStatus.Unsupported; item.Message = ex.Message; }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException ||
                    ex is System.Xml.XmlException || ex.GetType().Namespace.StartsWith("NPOI", StringComparison.Ordinal))
                { item.Status = CollectionImportStatus.Failed; item.Message = "Ўқиб бўлмади. Аввалги маълумот сақланди: " + ex.Message; }
                report.Files.Add(item);
                progress?.Report(new CollectionImportProgress { Path = path, CompletedFiles = index + 1, TotalFiles = paths.Length, ProcessedBytes = bytes });
            }
            Recount(next);
            // A larger collection may require a lower per-source retained-record budget.
            foreach (var source in next.Sources) Compact(source, sourceBudget);
            Recount(next);
            CollectionStore.Validate(next);
            token.ThrowIfCancellationRequested();
            return report;
        }

        public static PredictionSource BuildSource(string path, IEnumerable<string> paragraphs, int recordBudget = 20000,
            CancellationToken cancellation = default(CancellationToken))
        {
            var source = new PredictionSource { Path = path };
            var seen = new HashSet<string>(StringComparer.Ordinal);
            // Source identity includes all unique normalized sentences, even if the compact index prunes some.
            var allHashes = new List<string>();
            foreach (string paragraph in paragraphs)
            foreach (string sentence in Regex.Split(paragraph, @"[.!?;:\r\n\t\a\d()\[\]«»]+"))
            {
                cancellation.ThrowIfCancellationRequested();
                var words = TextHelper.Tokenize(sentence).Where(t => WordCompletionEngine.IsToken(t.Normalized)).ToArray();
                if (words.Length == 0) continue;
                string normalized = string.Join(" ", words.Select(w => w.Normalized));
                string hash = Hash(normalized);
                if (!seen.Add(hash)) continue;
                allHashes.Add(hash);
                source.WordCount += words.Length;
                var passage = new PredictionPassage { Hash = hash, WordCount = words.Length };
                for (int i = 0; i < words.Length; i++)
                {
                    if (!source.DisplayWords.ContainsKey(words[i].Normalized)) source.DisplayWords[words[i].Normalized] = words[i].Original;
                    var sequence = new StringBuilder();
                    for (int n = 0; n < 7 && i + n < words.Length; n++)
                    {
                        if (n > 0) sequence.Append(' ');
                        sequence.Append(words[i + n].Normalized);
                        string key = sequence.ToString(); int count;
                        passage.Sequences.TryGetValue(key, out count); passage.Sequences[key] = count + 1;
                    }
                }
                source.Passages.Add(passage);
                // Bound transient memory even for a document near the extraction limit.
                if (source.Passages.Sum(p => p.Sequences.Count) > 150000) Compact(source, Math.Max(recordBudget, 30000));
            }
            source.Hash = Hash(string.Join("\n", allHashes.OrderBy(h => h, StringComparer.Ordinal)));
            Compact(source, recordBudget);
            return source;
        }

        public static void Recount(PredictionCollection collection)
        {
            var seenPassages = new HashSet<string>(StringComparer.Ordinal);
            var seenSources = new HashSet<string>(StringComparer.Ordinal);
            foreach (var source in collection.Sources)
            {
                if (source.Passages.Count == 0) { if (!seenSources.Add(source.Hash)) source.Sequences.Clear(); continue; }
                seenSources.Add(source.Hash);
                source.Sequences = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (var passage in source.Passages)
                {
                    if (!seenPassages.Add(passage.Hash)) continue;
                    foreach (var pair in passage.Sequences) { int count; source.Sequences.TryGetValue(pair.Key, out count); source.Sequences[pair.Key] = count + pair.Value; }
                }
            }
        }

        private static void Compact(PredictionSource source, int budget)
        {
            if (source.Passages.Count == 0) return;
            var counts = source.Passages.SelectMany(p => p.Sequences).GroupBy(p => p.Key)
                .ToDictionary(g => g.Key, g => g.Sum(p => p.Value), StringComparer.Ordinal);
            // Retain repeated sequences first. Singletons remain available as vocabulary.
            var priority = counts.OrderByDescending(p => p.Value > 1).ThenByDescending(p => p.Value)
                .ThenBy(p => p.Key.Count(c => c == ' ')).ThenBy(p => p.Key, StringComparer.Ordinal).ToArray();
            var frequency = source.Passages.SelectMany(p => p.Sequences.Keys).GroupBy(k => k)
                .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);
            var keep = new HashSet<string>(StringComparer.Ordinal);
            int remaining = Math.Max(0, budget);
            foreach (var pair in priority)
            {
                int cost = frequency[pair.Key] * 2 + 2; // passage/count, effective count and optional display entry
                if (cost > remaining) continue;
                keep.Add(pair.Key); remaining -= cost;
            }
            foreach (var passage in source.Passages)
                passage.Sequences = passage.Sequences.Where(p => keep.Contains(p.Key)).ToDictionary(p => p.Key, p => p.Value, StringComparer.Ordinal);
            source.Passages.RemoveAll(p => p.Sequences.Count == 0);
            source.DisplayWords = source.DisplayWords.Where(p => keep.Contains(p.Key)).ToDictionary(p => p.Key, p => p.Value, StringComparer.Ordinal);
            source.Sequences = counts.Where(p => keep.Contains(p.Key)).ToDictionary(p => p.Key, p => p.Value, StringComparer.Ordinal);
        }
        private static PredictionSource Clone(PredictionSource source) => new PredictionSource
        {
            Path = source.Path, Hash = source.Hash, SourceUrl = source.SourceUrl, WordCount = source.WordCount,
            Sequences = new Dictionary<string, int>(source.Sequences, StringComparer.Ordinal),
            DisplayWords = new Dictionary<string, string>(source.DisplayWords, StringComparer.Ordinal),
            Passages = source.Passages.Select(p => new PredictionPassage { Hash = p.Hash, WordCount = p.WordCount,
                Sequences = new Dictionary<string, int>(p.Sequences, StringComparer.Ordinal) }).ToList()
        };
        public static string Hash(string value)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value))).Replace("-", "").ToLowerInvariant();
        }
    }
}
