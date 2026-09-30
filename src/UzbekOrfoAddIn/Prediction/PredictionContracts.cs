using System;
using System.Collections.Generic;
using System.Linq;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Prediction
{
    public sealed class PredictionRequest
    {
        public PredictionRequest()
            : this(string.Empty, string.Empty, ScriptType.Unknown, null, 0) { }

        public PredictionRequest(string precedingContext, string prefix, ScriptType script,
            IEnumerable<string> activeCollectionIds, long requestId, int maxResults = 3,
            IReadOnlyDictionary<string, int> acceptanceCounts = null)
        {
            PrecedingContext = precedingContext ?? string.Empty;
            Prefix = prefix ?? string.Empty;
            Script = script;
            ActiveCollectionIds = Array.AsReadOnly((activeCollectionIds ?? Enumerable.Empty<string>()).ToArray());
            RequestId = requestId;
            MaxResults = Math.Max(0, Math.Min(20, maxResults));
            AcceptanceCounts = acceptanceCounts;
        }

        public string PrecedingContext { get; set; }
        public string Prefix { get; set; }
        public ScriptType Script { get; set; }
        // An empty selection means no collection is enabled (including built-in collections).
        public IReadOnlyList<string> ActiveCollectionIds { get; set; }
        public long RequestId { get; set; }
        public int MaxResults { get; set; }
        // Key: collection ID + tab + TextHelper.NormalizeWord(fullCompletion).
        public IReadOnlyDictionary<string, int> AcceptanceCounts { get; set; }
    }

    public sealed class PredictionCandidate
    {
        internal PredictionCandidate(string tail, string fullCompletion, double score, int support,
            int contextWordCount, long requestId, IEnumerable<PredictionProvenance> provenance)
        {
            InsertableTail = tail;
            FullCompletion = fullCompletion;
            Score = score;
            Support = support;
            ContextWordCount = contextWordCount;
            RequestId = requestId;
            Provenance = Array.AsReadOnly(provenance.ToArray());
            CollectionId = Provenance.Count == 0 ? string.Empty : Provenance[0].CollectionId;
        }
        // Append only this text at the caret. FullCompletion includes the exact user-typed prefix.
        public string InsertableTail { get; private set; }
        public string FullCompletion { get; private set; }
        public string Completion { get { return FullCompletion; } }
        public string Continuation { get { return InsertableTail; } }
        public string CollectionId { get; private set; }
        public long RequestId { get; private set; }
        public double Score { get; private set; }
        public int Support { get; private set; }
        public int ContextWordCount { get; private set; }
        public IReadOnlyList<PredictionProvenance> Provenance { get; private set; }
    }

    public sealed class PredictionProvenance
    {
        internal PredictionProvenance(string collectionId, string collectionName, bool builtIn,
            string path, string hash, int support)
        {
            CollectionId = collectionId;
            CollectionName = collectionName;
            BuiltIn = builtIn;
            SourcePath = path;
            SourceHash = hash;
            Support = support;
        }
        public string CollectionId { get; private set; }
        public string CollectionName { get; private set; }
        public bool BuiltIn { get; private set; }
        public string SourcePath { get; private set; }
        public string SourceHash { get; private set; }
        public int Support { get; private set; }
    }
}
