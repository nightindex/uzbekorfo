using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Prediction
{
    internal enum PredictionRankingPolicy { Current, AccuracyCandidate }

    /// <summary>
    /// A detached, immutable corpus snapshot. Replace the engine when collections change.
    /// Acceptance counts come from a detached request snapshot and are collection-scoped.
    /// No morphology, stitching of continuations, or synthetic suffixes are used.
    /// </summary>
    public sealed class PhrasePredictionEngine : IPredictionEngine
    {
        private readonly Dictionary<string, Dictionary<string, Entry[]>> _collections;
        private readonly double _minimumProbability;
        private readonly double _minimumMargin;
        private readonly bool _selectiveNextWord;
        private readonly PredictionRankingPolicy _rankingPolicy;
        private readonly double _afterSpaceProbability;
        private readonly double _afterSpaceMargin;

        // Phrase thresholds: corpus-report.json. Additional after-space pilot
        // selectivity policy and its coverage tradeoff: docs/prediction/pilot.md.
        public PhrasePredictionEngine(IEnumerable<PredictionCollection> collections, double minimumProbability = 0.4, double minimumMargin = 0.0,
            bool selectiveNextWord = true)
            : this(collections, minimumProbability, minimumMargin, selectiveNextWord,
                PredictionRankingPolicy.Current, 0.4, 0.1) { }

        internal PhrasePredictionEngine(IEnumerable<PredictionCollection> collections, double minimumProbability,
            double minimumMargin, bool selectiveNextWord, PredictionRankingPolicy rankingPolicy,
            double afterSpaceProbability, double afterSpaceMargin)
        {
            _minimumProbability = minimumProbability;
            _minimumMargin = minimumMargin;
            _selectiveNextWord = selectiveNextWord;
            _rankingPolicy = rankingPolicy;
            _afterSpaceProbability = afterSpaceProbability;
            _afterSpaceMargin = afterSpaceMargin;
            _collections = new Dictionary<string, Dictionary<string, Entry[]>>(StringComparer.Ordinal);
            // Local to this snapshot: never intern private text in the process-wide
            // string pool, where it could survive removal of a collection.
            var strings = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var collection in collections ?? Enumerable.Empty<PredictionCollection>())
            {
                if (collection == null || string.IsNullOrWhiteSpace(collection.Id)) continue;
                if (_collections.ContainsKey(collection.Id))
                    throw new ArgumentException("Collection IDs must be unique.", "collections");
                var buckets = new Dictionary<string, Dictionary<string, Entry>>(StringComparer.Ordinal);
                var seenSources = new HashSet<string>(StringComparer.Ordinal);
                foreach (var source in collection.Sources ?? new List<PredictionSource>())
                {
                    if (source == null || source.Sequences == null) continue;
                    // Re-importing an identical source must not turn single evidence into repeated support.
                    var identity = string.IsNullOrEmpty(source.Hash) ? "path:" + source.Path : "hash:" + source.Hash;
                    if (!seenSources.Add(identity)) continue;
                    // Provenance is immutable and differs only by support within a
                    // source. Reuse it across contexts/completions with equal counts.
                    var provenanceBySupport = new Dictionary<int, PredictionProvenance>();
                    foreach (var sequence in source.Sequences)
                    {
                        if (sequence.Value <= 0 || string.IsNullOrWhiteSpace(sequence.Key)) continue;
                        var words = sequence.Key.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(TextHelper.NormalizeWord).ToArray();
                        if (words.Length == 0 || words.Length > 7 || words.Any(w => !IsToken(w))) continue;
                        for (int contextSize = 0; contextSize <= 2; contextSize++)
                        {
                            int wordCount = words.Length - contextSize;
                            if (wordCount < 1 || wordCount > 5 || (contextSize == 0 && wordCount != 1)) continue;
                            string context = Share(strings, string.Join(" ", words.Take(contextSize)));
                            string completion = Share(strings, string.Join(" ", words.Skip(contextSize)));
                            Dictionary<string, Entry> bucket;
                            if (!buckets.TryGetValue(context, out bucket))
                                buckets.Add(context, bucket = new Dictionary<string, Entry>(StringComparer.Ordinal));
                            Entry entry;
                            if (!bucket.TryGetValue(completion, out entry))
                                bucket.Add(completion, entry = new Entry(completion, wordCount));
                            PredictionProvenance provenance;
                            if (!provenanceBySupport.TryGetValue(sequence.Value, out provenance))
                            {
                                provenance = new PredictionProvenance(collection.Id, collection.Name,
                                    collection.BuiltIn, source.Path, source.Hash, sequence.Value);
                                provenanceBySupport.Add(sequence.Value, provenance);
                            }
                            entry.Provenance.Add(provenance);
                        }
                    }
                }
                _collections.Add(collection.Id, buckets.ToDictionary(b => b.Key,
                    b => b.Value.Values.OrderBy(e => e.Text, StringComparer.Ordinal).Select(e => e.Freeze()).ToArray(), StringComparer.Ordinal));
            }
        }

        private static string Share(Dictionary<string, string> strings, string value)
        {
            string existing;
            if (strings.TryGetValue(value, out existing)) return existing;
            strings.Add(value, value);
            return value;
        }

        public Task<PredictionCandidate[]> PredictAsync(PredictionRequest request, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            var snapshot = new PredictionRequest(request.PrecedingContext, request.Prefix, request.Script,
                request.ActiveCollectionIds, request.RequestId, request.MaxResults);
            snapshot.AcceptanceCounts = request.AcceptanceCounts == null ? null : request.AcceptanceCounts
                .Where(p => p.Key != null).Take(2048).ToDictionary(p => p.Key, p => Math.Max(0, Math.Min(20, p.Value)), StringComparer.Ordinal);
            return Task.Run(() => Predict(snapshot, cancellationToken), cancellationToken);
        }

        private PredictionCandidate[] Predict(PredictionRequest request, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            string prefix = TextHelper.NormalizeWord(request.Prefix);
            // The caret prefix is a single partial token, not punctuation or already completed words.
            if (request.MaxResults == 0 || (request.Prefix.Length > 0 &&
                (!IsToken(prefix) || prefix.Length != request.Prefix.Length)))
                return new PredictionCandidate[0];
            var script = request.Script;
            if (script == ScriptType.Unknown) script = DetectScript(prefix);
            if (script == ScriptType.Unknown) script = DetectScript(ContextWords(request.PrecedingContext).LastOrDefault() ?? string.Empty);
            if (script == ScriptType.Mixed) return new PredictionCandidate[0];

            var context = ContextWords(request.PrecedingContext);
            var candidates = new Dictionary<string, PredictionCandidate>(StringComparer.Ordinal);
            var selected = request.ActiveCollectionIds.Where(id => id != null).Distinct(StringComparer.Ordinal).ToArray();
            for (int size = Math.Min(2, context.Length); size >= 0; size--)
            {
                token.ThrowIfCancellationRequested();
                if (size == 0 && prefix.Length == 0) continue;
                string key = string.Join(" ", context.Skip(context.Length - size));
                var matched = new Dictionary<string, Match>(StringComparer.Ordinal);
                foreach (string id in selected)
                {
                    Dictionary<string, Entry[]> index;
                    Entry[] entries;
                    if (!_collections.TryGetValue(id, out index) || !index.TryGetValue(key, out entries)) continue;
                    for (int i = LowerBound(entries, prefix); i < entries.Length; i++)
                    {
                        token.ThrowIfCancellationRequested();
                        var entry = entries[i];
                        if (!entry.Text.StartsWith(prefix, StringComparison.Ordinal)) break;
                        if (entry.Text.Length <= prefix.Length || candidates.ContainsKey(entry.Text)) continue;
                        if (script != ScriptType.Unknown && DetectScript(entry.Text) != script) continue;
                        Match match;
                        if (!matched.TryGetValue(entry.Text, out match))
                            matched.Add(entry.Text, match = new Match(entry.WordCount));
                        match.Provenance.AddRange(entry.FrozenProvenance);
                    }
                }
                var nextWords = matched.Where(p => p.Value.WordCount == 1).ToDictionary(p => p.Key,
                    p => p.Value.Provenance.GroupBy(v => string.IsNullOrEmpty(v.SourceHash) ? v.SourcePath : v.SourceHash)
                        .Sum(g => (double)g.Max(v => v.Support)), StringComparer.Ordinal);
                double total = Math.Max(1, nextWords.Values.Sum());
                double best = nextWords.Values.DefaultIfEmpty(0).Max();
                double second = nextWords.Values.OrderByDescending(v => v).Skip(1).DefaultIfEmpty(0).First();
                foreach (var item in matched)
                {
                    // The same content in multiple enabled collections contributes evidence only once.
                    var evidence = item.Value.Provenance.GroupBy(p =>
                        string.IsNullOrEmpty(p.SourceHash) ? "path:" + p.SourcePath : "hash:" + p.SourceHash)
                        .Select(g => g.Max(p => p.Support));
                    int support = (int)Math.Min(int.MaxValue, evidence.Sum(n => (long)n));
                    if (support < (item.Value.WordCount > 1 || prefix.Length == 0 ? 2 : 1)) continue;
                    if (_selectiveNextWord && prefix.Length == 0)
                    {
                        // With no typed prefix, show only a supported leading next
                        // word. A tie or diffuse distribution should remain quiet.
                        // Apply the gate to phrases too, before any preference boost.
                        string firstWord = item.Key.Split(' ')[0];
                        double firstSupport;
                        if (!nextWords.TryGetValue(firstWord, out firstSupport) || firstSupport < 2 ||
                            firstSupport / total < _afterSpaceProbability ||
                            (firstSupport - second) / total < _afterSpaceMargin) continue;
                    }
                    if (item.Value.WordCount > 1 && (support / total < _minimumProbability || (best - second) / total < _minimumMargin)) continue;
                    double boost = 0;
                    if (request.AcceptanceCounts != null)
                    {
                        foreach (var id in item.Value.Provenance.Select(p => p.CollectionId).Distinct())
                        {
                            int accepted;
                            if (request.AcceptanceCounts.TryGetValue(id + "\t" + item.Key, out accepted))
                                boost = Math.Max(boost, Math.Min(5.0, accepted * 0.25));
                        }
                    }
                    string tail = item.Key.Substring(prefix.Length);
                    if (request.Prefix.Count(char.IsLetter) > 1 && request.Prefix.Where(char.IsLetter).All(char.IsUpper)) tail = tail.ToUpperInvariant();
                    string display = PredictionCasing.Apply(request.PrecedingContext, request.Prefix, request.Prefix + tail);
                    tail = display.Substring(request.Prefix.Length);
                    double probability = 0;
                    if (_rankingPolicy == PredictionRankingPolicy.AccuracyCandidate)
                    {
                        double firstSupportForScore;
                        nextWords.TryGetValue(item.Key.Split(' ')[0], out firstSupportForScore);
                        probability = firstSupportForScore / total;
                    }
                    double score = _rankingPolicy == PredictionRankingPolicy.AccuracyCandidate
                        ? (prefix.Length == 0 ? probability * 100 + size * 8 + Math.Log(1.0 + support) * 5
                            : probability * 50 + size * 30 + Math.Log(1.0 + support) * 2)
                            - (item.Value.WordCount - 1) * 3 + boost
                        : size * 100 + Math.Min(20, Math.Log(1.0 + support) * 2)
                            + item.Value.WordCount + boost;
                    candidates.Add(item.Key, new PredictionCandidate(tail, request.Prefix + tail, score, support, size, request.RequestId,
                        item.Value.Provenance.OrderBy(p => p.CollectionId, StringComparer.Ordinal)
                            .ThenBy(p => p.SourcePath, StringComparer.Ordinal)));
                }
            }
            return candidates.Values.OrderByDescending(c => c.Score)
                .ThenBy(c => c.FullCompletion, StringComparer.Ordinal).Take(request.MaxResults).ToArray();
        }

        private static int LowerBound(Entry[] entries, string prefix)
        {
            int low = 0, high = entries.Length;
            while (low < high)
            {
                int mid = low + (high - low) / 2;
                if (string.CompareOrdinal(entries[mid].Text, prefix) < 0) low = mid + 1;
                else high = mid;
            }
            return low;
        }

        private static string[] ContextWords(string context)
        {
            // Never use context from a previous sentence, paragraph, or table cell.
            int boundary = context.LastIndexOfAny(new[] { '.', '!', '?', ';', ':', '\r', '\n', '\t', '\a' });
            return TextHelper.Tokenize(context.Substring(boundary + 1)).Select(t => t.Normalized).Reverse().Take(2).Reverse().ToArray();
        }

        private static bool IsToken(string value)
        {
            if (string.IsNullOrEmpty(value) || !value.Any(char.IsLetter)) return false;
            return value.All(c => char.IsLetter(c) || c == '\'' || c == '-');
        }

        private static ScriptType DetectScript(string value)
        {
            bool latin = value.Any(TextHelper.IsLatin), cyrillic = value.Any(TextHelper.IsCyrillic);
            return latin && cyrillic ? ScriptType.Mixed : latin ? ScriptType.Latin : cyrillic ? ScriptType.Cyrillic : ScriptType.Unknown;
        }

        private sealed class Entry
        {
            internal readonly string Text;
            internal readonly int WordCount;
            internal List<PredictionProvenance> Provenance = new List<PredictionProvenance>();
            internal PredictionProvenance[] FrozenProvenance;
            internal Entry(string text, int wordCount) { Text = text; WordCount = wordCount; }
            internal Entry Freeze()
            {
                FrozenProvenance = Provenance.ToArray();
                Provenance = null; // Release list overhead and unused backing capacity.
                return this;
            }
        }

        private sealed class Match
        {
            internal readonly int WordCount;
            internal readonly List<PredictionProvenance> Provenance = new List<PredictionProvenance>();
            internal Match(int wordCount) { WordCount = wordCount; }
        }
    }
}
