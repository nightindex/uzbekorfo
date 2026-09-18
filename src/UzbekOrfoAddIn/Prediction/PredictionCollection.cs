using System;
using System.Collections.Generic;

namespace UzbekOrfoAddIn.Prediction
{
    public sealed class PredictionCollection
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Name { get; set; } = string.Empty;
        public bool BuiltIn { get; set; }
        public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
        public List<PredictionSource> Sources { get; set; } = new List<PredictionSource>();
        public Dictionary<string, bool> ImportLocations { get; set; } = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        public List<string> ExcludedSources { get; set; } = new List<string>();
    }

    public sealed class PredictionSource
    {
        public string Path { get; set; } = string.Empty;
        public string SourceUrl { get; set; }
        /// <summary>SHA-256 of the ordered, normalized, unique passage hashes.</summary>
        public string Hash { get; set; } = string.Empty;
        /// <summary>Observed words, excluding repeated passages within this source.</summary>
        public int WordCount { get; set; }
        /// <summary>
        /// Effective collection contributions: normalized 1..7-word keys. A passage
        /// contributes through only the first source that contains it, in Sources order.
        /// </summary>
        public Dictionary<string, int> Sequences { get; set; } = new Dictionary<string, int>(StringComparer.Ordinal);
        public Dictionary<string, string> DisplayWords { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
        /// <summary>
        /// Hashed passages and their counts, never the original passage text. Retained
        /// even for duplicates so ownership can transfer when another source changes.
        /// An empty list also supports legacy/built-in sources with Sequences only.
        /// </summary>
        public List<PredictionPassage> Passages { get; set; } = new List<PredictionPassage>();
    }

    public sealed class PredictionPassage
    {
        public string Hash { get; set; } = string.Empty;
        public int WordCount { get; set; }
        public Dictionary<string, int> Sequences { get; set; } = new Dictionary<string, int>(StringComparer.Ordinal);
    }

    public sealed class ExtractedDocument
    {
        public List<string> Passages { get; set; } = new List<string>();
    }

    public enum CollectionImportStatus { Imported, Duplicate, Missing, Failed, LimitExceeded, Unsupported }

    public sealed class CollectionImportFileReport
    {
        public string Path { get; set; }
        public CollectionImportStatus Status { get; set; }
        public string Message { get; set; }
        public int WordCount { get; set; }
    }

    public sealed class CollectionImportProgress
    {
        public string Path { get; set; }
        public int CompletedFiles { get; set; }
        public int TotalFiles { get; set; }
        public long ProcessedBytes { get; set; }
    }

    public sealed class CollectionImportResult
    {
        public PredictionCollection Collection { get; set; }
        public List<CollectionImportFileReport> Files { get; set; } = new List<CollectionImportFileReport>();
    }
}
