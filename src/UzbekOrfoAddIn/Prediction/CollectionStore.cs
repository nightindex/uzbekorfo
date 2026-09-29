using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Newtonsoft.Json;

namespace UzbekOrfoAddIn.Prediction
{
    /// <summary>
    /// Versioned per-collection JSON. Private collections are gzip then CurrentUser
    /// DPAPI; built-ins are plain gzip JSON. No unencrypted private temporary files.
    /// </summary>
    public sealed class CollectionStore
    {
        public const int FormatVersion = 1;
        public const int MaxModelRecords = 200000;
        private const int MaxStoredBytes = 64 * 1024 * 1024;
        private const int MaxJsonBytes = 128 * 1024 * 1024;
        private const string PrivateSuffix = ".private.gz.dpapi";
        private const string BuiltInSuffix = ".builtin.json.gz";
        private static readonly object Gate = new object();
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("UzbekOrfo.MatnAI.Collection.v1");
        private static readonly Regex SafeId = new Regex(@"\A[a-zA-Z0-9][a-zA-Z0-9_-]{0,99}\z", RegexOptions.CultureInvariant);
        public string Root { get; private set; }
        public string RootDirectory => Root;
        /// <summary>Errors and backup recoveries from the last LoadAll; healthy collections still load.</summary>
        public List<string> LoadErrors { get; private set; } = new List<string>();

        public CollectionStore(string root = null)
        {
            Root = Path.GetFullPath(root ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "UzbekOrfo", "MatnAI"));
        }

        public List<PredictionCollection> LoadAll()
        {
            lock (Gate)
            {
                LoadErrors.Clear();
                var result = new List<PredictionCollection>();
                foreach (string resource in new[] { "legal_prediction.collection.gz", "legal_prediction_latin.collection.gz" })
                using (var embedded = typeof(CollectionStore).Assembly.GetManifestResourceStream("UzbekOrfoAddIn.Data." + resource))
                {
                    if (embedded != null)
                    {
                        using (var gzip = new GZipStream(embedded, CompressionMode.Decompress))
                        using (var reader = new StreamReader(new BoundedStream(gzip, MaxJsonBytes), Encoding.UTF8))
                        using (var json = new JsonTextReader(reader) { MaxDepth = 32 })
                        {
                            var data = CreateSerializer().Deserialize<Envelope>(json);
                            if (data == null || data.Version != FormatVersion || data.Collection == null || !data.Collection.BuiltIn)
                                throw new InvalidDataException("Invalid embedded legal collection.");
                            Validate(data.Collection); result.Add(data.Collection);
                        }
                    }
                }
                if (!Directory.Exists(Root)) return result;
                var paths = Directory.EnumerateFiles(Root).Select(p => p.EndsWith(".bak", StringComparison.OrdinalIgnoreCase) ? p.Substring(0, p.Length - 4) : p)
                    .Where(IsCollectionPath).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(p => p, StringComparer.Ordinal).ToList();
                var ids = new HashSet<string>(result.Select(c => c.Id), StringComparer.OrdinalIgnoreCase);
                foreach (string path in paths)
                {
                    PredictionCollection collection;
                    try { collection = Read(path); }
                    catch (Exception error) when (IsReadError(error))
                    {
                        try
                        {
                            collection = Read(path + ".bak", path);
                            LoadErrors.Add(Path.GetFileName(path) + ": loaded backup after " + error.Message);
                        }
                        catch (Exception backupError) when (IsReadError(backupError))
                        {
                            LoadErrors.Add(Path.GetFileName(path) + ": " + error.Message + " Backup: " + backupError.Message);
                            continue;
                        }
                    }
                    if (ids.Add(collection.Id)) result.Add(collection);
                    else LoadErrors.Add("Duplicate collection ID: " + collection.Id);
                }
                return result;
            }
        }

        public void Save(PredictionCollection collection, bool discardPrevious = false, CancellationToken cancellation = default(CancellationToken))
        {
            cancellation.ThrowIfCancellationRequested();
            Validate(collection);
            lock (Gate)
            {
                Directory.CreateDirectory(Root);
                using (AcquireWriteLock())
                {
                    string path = GetPath(collection.Id, collection.BuiltIn);
                    string opposite = GetPath(collection.Id, !collection.BuiltIn);
                    if (File.Exists(opposite) || File.Exists(opposite + ".bak"))
                        throw new InvalidOperationException("Cannot change an existing collection's BuiltIn flag.");
                    bool replaceBackup = true;
                    if (File.Exists(path))
                    {
                        try { Read(path); }
                        catch (NotSupportedException) { throw; } // Do not overwrite a newer format.
                        catch (Exception error) when (IsReadError(error))
                        {
                            // Keep the last good backup when Save follows fallback recovery.
                            Read(path + ".bak", path);
                            replaceBackup = false;
                        }
                    }
                    byte[] bytes;
                    using (var compressed = new MemoryStream())
                    {
                        using (var gzip = new GZipStream(compressed, CompressionLevel.Optimal, true))
                        using (var bounded = new BoundedStream(gzip, MaxJsonBytes))
                        using (var writer = new StreamWriter(bounded, new UTF8Encoding(false)))
                        using (var json = new JsonTextWriter(writer))
                            CreateSerializer().Serialize(json, new Envelope { Version = FormatVersion, Collection = collection });
                        bytes = compressed.ToArray();
                    }
                    if (!collection.BuiltIn) bytes = ProtectedData.Protect(bytes, Entropy, DataProtectionScope.CurrentUser);
                    if (bytes.Length > MaxStoredBytes) throw new InvalidDataException("Collection exceeds the stored-size limit.");
                    string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                    try
                    {
                        using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                        {
                            stream.Write(bytes, 0, bytes.Length);
                            stream.Flush(true);
                        }
                        // Publication is the commit point. A cancelled import must
                        // leave the previous usable model in place.
                        cancellation.ThrowIfCancellationRequested();
                        if (File.Exists(path)) File.Replace(temporary, path, replaceBackup ? path + ".bak" : null);
                        else File.Move(temporary, path);
                        if (discardPrevious && File.Exists(path + ".bak")) File.Delete(path + ".bak");
                    }
                    finally { if (File.Exists(temporary)) File.Delete(temporary); }
                }
            }
        }

        /// <summary>Deletes private model files and their backup only; never deletes source documents.</summary>
        public void Delete(string id)
        {
            CheckId(id);
            lock (Gate)
            {
                if (!Directory.Exists(Root)) return;
                using (AcquireWriteLock())
                {
                    if (File.Exists(GetPath(id, true)) || File.Exists(GetPath(id, true) + ".bak"))
                        throw new InvalidOperationException("Built-in collections cannot be deleted.");
                    string path = GetPath(id, false);
                    foreach (string temporary in Directory.EnumerateFiles(Root, id + PrivateSuffix + ".*.tmp"))
                    {
                        string suffix = temporary.Substring(path.Length + 1);
                        Guid stagingId;
                        if (Guid.TryParseExact(suffix.Substring(0, suffix.Length - 4), "N", out stagingId)) File.Delete(temporary);
                    }
                    // Backup first: interruption must not resurrect an explicitly deleted collection.
                    if (File.Exists(path + ".bak")) File.Delete(path + ".bak");
                    if (File.Exists(path)) File.Delete(path);
                }
            }
        }

        private PredictionCollection Read(string path, string originalPath = null)
        {
            string original = originalPath ?? path;
            bool builtIn = original.EndsWith(BuiltInSuffix, StringComparison.OrdinalIgnoreCase);
            byte[] bytes;
            using (var stream = File.OpenRead(path))
            {
                if (stream.Length > MaxStoredBytes) throw new InvalidDataException("Stored collection exceeds 64 MiB.");
                using (var memory = new MemoryStream()) { stream.CopyTo(memory); bytes = memory.ToArray(); }
            }
            if (!builtIn) bytes = ProtectedData.Unprotect(bytes, Entropy, DataProtectionScope.CurrentUser);
            using (var memory = new MemoryStream(bytes, false))
            using (var gzip = new GZipStream(memory, CompressionMode.Decompress))
            using (var bounded = new BoundedStream(gzip, MaxJsonBytes))
            using (var reader = new StreamReader(bounded, new UTF8Encoding(false, true)))
            using (var json = new JsonTextReader(reader) { MaxDepth = 32 })
            {
                var envelope = CreateSerializer().Deserialize<Envelope>(json);
                if (envelope == null) throw new InvalidDataException("Empty collection.");
                if (envelope.Version != FormatVersion) throw new NotSupportedException("Unsupported collection format " + envelope.Version + ".");
                if (json.Read()) throw new InvalidDataException("Trailing JSON content.");
                Validate(envelope.Collection);
                var collection = envelope.Collection;
                if (collection.BuiltIn != builtIn || !string.Equals(GetPath(collection.Id, builtIn), original, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Collection identity does not match its file.");
                return collection;
            }
        }

        private FileStream AcquireWriteLock()
        {
            // An overlapping process fails safely; callers can retry without torn writes.
            return new FileStream(Path.Combine(Root, ".collection-write.lock"), FileMode.OpenOrCreate, FileAccess.Write, FileShare.None);
        }
        private string GetPath(string id, bool builtIn) { CheckId(id); return Path.Combine(Root, id + (builtIn ? BuiltInSuffix : PrivateSuffix)); }
        private static bool IsCollectionPath(string path) { return path.EndsWith(PrivateSuffix, StringComparison.OrdinalIgnoreCase) || path.EndsWith(BuiltInSuffix, StringComparison.OrdinalIgnoreCase); }
        private static void CheckId(string id)
        {
            if (id == null || !SafeId.IsMatch(id)) throw new ArgumentException("Collection IDs must contain 1..100 ASCII letters, digits, underscores or hyphens, starting with a letter or digit.", "id");
            string first = id.Split('.')[0];
            if (Regex.IsMatch(first, @"\A(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])\z", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                throw new ArgumentException("Reserved collection ID.", "id");
        }
        private static JsonSerializer CreateSerializer()
        {
            return JsonSerializer.Create(new JsonSerializerSettings
            {
                TypeNameHandling = TypeNameHandling.None, MaxDepth = 32,
                DateTimeZoneHandling = DateTimeZoneHandling.Utc, CheckAdditionalContent = true
            });
        }
        private static bool IsReadError(Exception error)
        {
            return error is IOException || error is InvalidDataException || error is UnauthorizedAccessException || error is CryptographicException ||
                error is JsonException || error is ArgumentException || error is NotSupportedException;
        }

        internal static void Validate(PredictionCollection collection)
        {
            if (collection == null) throw new ArgumentNullException("collection");
            CheckId(collection.Id);
            if (collection.Name == null || collection.Name.Length > 1000 || collection.Sources == null || collection.Sources.Count > CollectionImporter.MaxFiles)
                throw new InvalidDataException("Invalid collection metadata.");
            if (collection.ImportLocations == null || collection.ExcludedSources == null ||
                collection.ImportLocations.Count > 1000 || collection.ExcludedSources.Count > 1000 ||
                collection.ImportLocations.Keys.Concat(collection.ExcludedSources).Any(p => string.IsNullOrEmpty(p) || p.Length > 32767))
                throw new InvalidDataException("Invalid import locations.");
            long records = 0;
            foreach (var source in collection.Sources)
            {
                if (source == null || source.Path == null || source.Path.Length > 32767 || source.Hash == null || source.Hash.Length > 128 ||
                    source.WordCount < 0 || (source.SourceUrl != null && source.SourceUrl.Length > 8192) ||
                    source.Passages == null || source.DisplayWords == null)
                    throw new InvalidDataException("Invalid source metadata.");
                records += ValidateSequences(source.Sequences);
                records += source.DisplayWords.Count + source.Passages.Count;
                foreach (var display in source.DisplayWords)
                    if (string.IsNullOrEmpty(display.Key) || display.Key.Length > 128 || string.IsNullOrEmpty(display.Value) || display.Value.Length > 128)
                        throw new InvalidDataException("Invalid display spelling.");
                foreach (var passage in source.Passages)
                {
                    if (passage == null || string.IsNullOrEmpty(passage.Hash) || passage.Hash.Length > 128 || passage.WordCount < 0)
                        throw new InvalidDataException("Invalid passage metadata.");
                    records += ValidateSequences(passage.Sequences);
                }
                if (records > MaxModelRecords) throw new InvalidDataException("Collection exceeds 200,000 model records.");
            }
        }
        private static int ValidateSequences(Dictionary<string, int> sequences)
        {
            if (sequences == null) throw new InvalidDataException("Missing sequence counts.");
            foreach (var pair in sequences)
                if (string.IsNullOrWhiteSpace(pair.Key) || pair.Key.Length > 902 || pair.Value <= 0 ||
                    pair.Key[0] == ' ' || pair.Key[pair.Key.Length - 1] == ' ' || pair.Key.Contains("  ") || pair.Key.Count(c => c == ' ') > 6)
                    throw new InvalidDataException("Invalid sequence count.");
            return sequences.Count;
        }

        private sealed class Envelope
        {
            public int Version { get; set; }
            public PredictionCollection Collection { get; set; }
        }

        private sealed class BoundedStream : Stream
        {
            private readonly Stream inner;
            private readonly long limit;
            private long total;
            internal BoundedStream(Stream inner, long limit) { this.inner = inner; this.limit = limit; }
            public override int Read(byte[] buffer, int offset, int count)
            {
                int read = inner.Read(buffer, offset, (int)Math.Min(count, limit - total + 1));
                Add(read); return read;
            }
            public override void Write(byte[] buffer, int offset, int count) { Add(count); inner.Write(buffer, offset, count); }
            private void Add(int count) { total += count; if (total > limit) throw new InvalidDataException("Collection JSON exceeds 128 MiB."); }
            public override void Flush() { inner.Flush(); }
            public override bool CanRead { get { return inner.CanRead; } }
            public override bool CanWrite { get { return inner.CanWrite; } }
            public override bool CanSeek { get { return false; } }
            public override long Length { get { throw new NotSupportedException(); } }
            public override long Position { get { return total; } set { throw new NotSupportedException(); } }
            public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
            public override void SetLength(long value) { throw new NotSupportedException(); }
        }
    }
}
