using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Maintains a private, local frequency report of words rejected during an
    /// explicit document check. No telemetry or document context leaves the PC.
    /// </summary>
    public sealed class UnknownWordReportService
    {
        private const int MaxUniqueWords = 10000;
        private const string Header = "word\tcount\tfirst_seen_utc\tlast_seen_utc";
        private readonly string _path;
        private readonly object _sync = new object();
        private readonly Dictionary<string, UnknownWordStat> _entries =
            new Dictionary<string, UnknownWordStat>(StringComparer.OrdinalIgnoreCase);

        public UnknownWordReportService(string path)
        {
            _path = path ?? throw new ArgumentNullException(nameof(path));
            Load();
        }

        public void RecordBatch(IEnumerable<string> words)
        {
            if (words == null) return;
            DateTime now = DateTime.UtcNow;
            bool changed = false;

            lock (_sync)
            {
                foreach (string raw in words)
                {
                    string word = TextHelper.NormalizeWord(raw);
                    if (string.IsNullOrWhiteSpace(word) || TextHelper.ShouldSkipWord(word)) continue;

                    if (_entries.TryGetValue(word, out UnknownWordStat existing))
                    {
                        existing.Count++;
                        existing.LastSeenUtc = now;
                        changed = true;
                    }
                    else if (_entries.Count < MaxUniqueWords)
                    {
                        _entries[word] = new UnknownWordStat
                        {
                            Word = word,
                            Count = 1,
                            FirstSeenUtc = now,
                            LastSeenUtc = now
                        };
                        changed = true;
                    }
                }

                if (changed) SaveLocked();
            }
        }

        public List<UnknownWordStat> GetSnapshot()
        {
            lock (_sync)
            {
                return _entries.Values
                    .OrderByDescending(e => e.Count)
                    .ThenBy(e => e.Word, StringComparer.OrdinalIgnoreCase)
                    .Select(Clone)
                    .ToList();
            }
        }

        public int ExportCsv(string path)
        {
            List<UnknownWordStat> snapshot = GetSnapshot();
            var lines = new List<string> { "word,count,first_seen_utc,last_seen_utc" };
            lines.AddRange(snapshot.Select(e => string.Join(",",
                Csv(e.Word),
                e.Count.ToString(CultureInfo.InvariantCulture),
                e.FirstSeenUtc.ToString("o", CultureInfo.InvariantCulture),
                e.LastSeenUtc.ToString("o", CultureInfo.InvariantCulture))));
            AtomicFile.WriteAllLines(path, lines.ToArray());
            return snapshot.Count;
        }

        private void Load()
        {
            if (!File.Exists(_path)) return;
            try
            {
                foreach (string line in File.ReadLines(_path).Skip(1))
                {
                    string[] parts = line.Split('\t');
                    if (parts.Length != 4 || !int.TryParse(parts[1], out int count) || count < 1 ||
                        !DateTime.TryParse(parts[2], CultureInfo.InvariantCulture,
                            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out DateTime first) ||
                        !DateTime.TryParse(parts[3], CultureInfo.InvariantCulture,
                            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out DateTime last))
                        continue;

                    string word = TextHelper.NormalizeWord(parts[0]);
                    if (string.IsNullOrWhiteSpace(word)) continue;
                    _entries[word] = new UnknownWordStat
                    {
                        Word = word,
                        Count = count,
                        FirstSeenUtc = first,
                        LastSeenUtc = last
                    };
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to load the local unknown-word report", ex);
                _entries.Clear();
            }
        }

        private void SaveLocked()
        {
            var lines = new List<string> { Header };
            lines.AddRange(_entries.Values
                .OrderByDescending(e => e.Count)
                .ThenBy(e => e.Word, StringComparer.OrdinalIgnoreCase)
                .Select(e => string.Join("\t",
                    e.Word.Replace("\t", string.Empty),
                    e.Count.ToString(CultureInfo.InvariantCulture),
                    e.FirstSeenUtc.ToString("o", CultureInfo.InvariantCulture),
                    e.LastSeenUtc.ToString("o", CultureInfo.InvariantCulture))));
            AtomicFile.WriteAllLines(_path, lines.ToArray());
        }

        private static UnknownWordStat Clone(UnknownWordStat source)
        {
            return new UnknownWordStat
            {
                Word = source.Word,
                Count = source.Count,
                FirstSeenUtc = source.FirstSeenUtc,
                LastSeenUtc = source.LastSeenUtc
            };
        }

        private static string Csv(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\"", "\"\"") + "\"";
        }
    }
}
