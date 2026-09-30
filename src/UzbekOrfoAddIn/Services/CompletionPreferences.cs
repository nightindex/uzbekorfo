using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>Acceptance counts only; no preceding context or document text. Caller owns consent.</summary>
    public sealed class CompletionPreferences
    {
        private readonly string _path;
        private readonly Dictionary<string, int> _counts = new Dictionary<string, int>(StringComparer.Ordinal);
        public CompletionPreferences(string path) { _path = path; }
        public Dictionary<string, int> Snapshot() => new Dictionary<string, int>(_counts, StringComparer.Ordinal);
        public void Load()
        {
            _counts.Clear();
            if (!File.Exists(_path) || new FileInfo(_path).Length > 1024 * 1024) return;
            var rows = File.ReadAllLines(_path);
            if (rows.Length == 0 || rows[0] != "matnai-acceptances-v1") return;
            foreach (var line in rows.Skip(1).Take(1000))
            {
                var fields = line.Split('\t');
                int count;
                if (fields.Length == 2 && WordCompletionEngine.IsToken(fields[0]) &&
                    int.TryParse(fields[1], NumberStyles.None, CultureInfo.InvariantCulture, out count) && count > 0 && count <= 1000000)
                    _counts[TextHelper.NormalizeWord(fields[0])] = count;
            }
        }
        public void Record(string accepted)
        {
            if (!WordCompletionEngine.IsToken(accepted)) return;
            string key = TextHelper.NormalizeWord(accepted);
            int previous;
            if (!_counts.TryGetValue(key, out previous) && _counts.Count >= 1000) return;
            _counts[key] = Math.Min(previous + 1, 1000000);
        }
        public void Save() => AtomicFile.WriteAllLines(_path, new[] { "matnai-acceptances-v1" }
            .Concat(_counts.OrderBy(p => p.Key, StringComparer.Ordinal)
                .Select(p => p.Key + "\t" + p.Value.ToString(CultureInfo.InvariantCulture))).ToArray());
        public void Clear()
        {
            _counts.Clear();
            Save(); // Only this model is reset; never a dictionary or general settings file.
        }
    }
}
