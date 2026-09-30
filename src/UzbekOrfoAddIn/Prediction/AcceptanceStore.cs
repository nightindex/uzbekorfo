using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Services;

namespace UzbekOrfoAddIn.Prediction
{
    /// <summary>Versioned, user-encrypted counts. No surrounding context is persisted.</summary>
    public sealed class AcceptanceStore
    {
        public const string DictionaryScope = "dictionary";
        private readonly string _path;
        private readonly Dictionary<string, int> _counts = new Dictionary<string, int>(StringComparer.Ordinal);
        public AcceptanceStore(string directory) { _path = Path.Combine(directory, "acceptances-v2.bin"); }
        public void Load(string legacyPath)
        {
            _counts.Clear();
            if (File.Exists(_path))
            {
                if (new FileInfo(_path).Length > 4 * 1024 * 1024) throw new InvalidDataException("Acceptance model too large.");
                var text = Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(_path), null, DataProtectionScope.CurrentUser));
                var lines = text.Split('\n');
                if (lines[0] != "matnai-acceptances-v2") throw new InvalidDataException("Unknown acceptance schema.");
                foreach (var line in lines.Skip(1).Take(10000))
                {
                    var fields = line.Split('\t'); int count;
                    if (fields.Length == 3 && int.TryParse(fields[2], out count) && count > 0 && Valid(fields[1]))
                        _counts[Key(fields[0], fields[1])] = Math.Min(count, 1000000);
                }
            }
            else if (File.Exists(legacyPath))
            {
                var legacy = new CompletionPreferences(legacyPath); legacy.Load();
                foreach (var entry in legacy.Snapshot()) _counts[Key(DictionaryScope, entry.Key)] = entry.Value;
                Save();
            }
        }
        public static string Key(string collectionId, string completion) => collectionId + "\t" + TextHelper.NormalizeWord(completion);
        private static bool Valid(string completion) => !string.IsNullOrWhiteSpace(completion) &&
            completion.Split(' ').Length <= 5 && completion.Split(' ').All(WordCompletionEngine.IsToken);
        public Dictionary<string, int> Snapshot() => new Dictionary<string, int>(_counts, StringComparer.Ordinal);
        public Dictionary<string, int> DictionarySnapshot() => _counts.Where(p => p.Key.StartsWith(DictionaryScope + "\t", StringComparison.Ordinal))
            .ToDictionary(p => p.Key.Substring(DictionaryScope.Length + 1), p => p.Value, StringComparer.Ordinal);
        public void Record(string collectionId, string completion)
        {
            if (string.IsNullOrEmpty(collectionId) || collectionId.Contains("\t") || !Valid(completion)) return;
            string key = Key(collectionId, completion); int count;
            if (!_counts.TryGetValue(key, out count) && _counts.Count >= 10000) return;
            _counts[key] = Math.Min(count + 1, 1000000); Save();
        }
        public void RemoveMissingCollections(IEnumerable<string> ids)
        {
            var allowed = new HashSet<string>(ids, StringComparer.Ordinal) { DictionaryScope };
            var removed = _counts.Keys.Where(k => !allowed.Contains(k.Split('\t')[0])).ToArray();
            foreach (string key in removed) _counts.Remove(key);
            if (removed.Length > 0) Save();
        }
        public void RemoveCollection(string id)
        {
            foreach (var key in _counts.Keys.Where(k => k.StartsWith(id + "\t", StringComparison.Ordinal)).ToArray()) _counts.Remove(key);
            Save();
        }
        public void Clear(string legacyPath)
        {
            _counts.Clear(); Save(); // Keep the v2 marker to prevent reimporting old counts.
            foreach (var path in new[] { legacyPath, legacyPath + ".bak", _path + ".bak" })
                if (File.Exists(path)) File.Delete(path);
        }
        private void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_path));
            string text = "matnai-acceptances-v2\n" + string.Join("\n", _counts.OrderBy(p => p.Key, StringComparer.Ordinal)
                .Select(p => p.Key + "\t" + p.Value.ToString(CultureInfo.InvariantCulture)));
            byte[] bytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(text), null, DataProtectionScope.CurrentUser);
            string temporary = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllBytes(temporary, bytes);
                if (File.Exists(_path)) File.Replace(temporary, _path, null); else File.Move(temporary, _path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
