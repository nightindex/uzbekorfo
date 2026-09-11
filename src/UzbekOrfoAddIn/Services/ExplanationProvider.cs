using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Provides word definitions, spelling rules, grammar notes, and usage examples
    /// for Uzbek words. Loaded from a JSON database file.
    /// Implements IExplanationProvider.
    /// </summary>
    public class ExplanationProvider : IExplanationProvider
    {
        private readonly string _dataPath;
        private readonly string _bundledDictionaryPath;
        private readonly Action _ensureBundledDictionary;
        private readonly object _bundledLoadLock = new object();
        private Dictionary<string, ExplanationEntry> _entries;
        private Dictionary<string, ExplanationEntry> _bundledEntries;
        private volatile bool _bundledEntriesLoaded;

        private sealed class BundledDictionaryPayload
        {
            public List<ExplanationEntry> Entries { get; set; }
        }

        // =====================================================================
        //  CONSTRUCTOR
        // =====================================================================

        public ExplanationProvider(string dataPath)
            : this(dataPath, null, null)
        {
        }

        /// <summary>
        /// Creates an explanation provider with an optional structured
        /// dictionary metadata fallback. The fallback is seeded and parsed only after
        /// a local explanation lookup misses.
        /// </summary>
        public ExplanationProvider(
            string dataPath,
            string bundledDictionaryPath,
            Action ensureBundledDictionary)
        {
            _dataPath = dataPath;
            _bundledDictionaryPath = bundledDictionaryPath;
            _ensureBundledDictionary = ensureBundledDictionary;
            _entries = new Dictionary<string, ExplanationEntry>(StringComparer.OrdinalIgnoreCase);
            _bundledEntries = new Dictionary<string, ExplanationEntry>(StringComparer.OrdinalIgnoreCase);
        }

        // =====================================================================
        //  IExplanationProvider
        // =====================================================================

        public int EntryCount => _entries.Count;

        /// <summary>
        /// Returns all explanation entries as a detached snapshot.
        /// </summary>
        public List<ExplanationEntry> GetAllEntries()
        {
            var list = new List<ExplanationEntry>(_entries.Count);

            foreach (var entry in _entries.Values)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.Word)) continue;

                list.Add(new ExplanationEntry
                {
                    Word = entry.Word,
                    Definition = entry.Definition,
                    SpellingRule = entry.SpellingRule,
                    GrammarNote = entry.GrammarNote,
                    Examples = entry.Examples != null ? entry.Examples.ToArray() : null
                });
            }

            list.Sort((a, b) => string.Compare(a.Word, b.Word, StringComparison.OrdinalIgnoreCase));
            return list;
        }

        /// <summary>
        /// Returns the explanation for a word. Tries exact match first,
        /// then normalized (lowercase, trimmed), then stem-based fuzzy match.
        /// </summary>
        public ExplanationEntry GetExplanation(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return null;

            string normalized = TextHelper.NormalizeWord(word);

            // 1. Local explanations override the bundled dataset.
            ExplanationEntry entry = FindExactOrStem(_entries, normalized);
            if (entry != null) return entry;

            // 2. Keep the existing fuzzy behavior for the small editable
            // local store. A fuzzy scan of the full bundled dataset would be
            // unnecessarily expensive on the UI thread.
            foreach (var kvp in _entries)
            {
                if (TextHelper.EditDistance(normalized, kvp.Key) <= 1)
                    return kvp.Value;
            }

            // 3. The generated metadata index contains built-in metadata. It is
            // intentionally loaded only after the local store cannot answer.
            EnsureBundledEntriesLoaded();
            return FindExactOrStem(_bundledEntries, normalized);
        }

        /// <summary>
        /// Checks if an explanation exists (exact or stem match).
        /// </summary>
        public bool HasExplanation(string word)
        {
            return GetExplanation(word) != null;
        }

        /// <summary>
        /// Adds or updates an explanation entry for a word.
        /// </summary>
        public void AddOrUpdate(ExplanationEntry entry)
        {
            if (entry == null || string.IsNullOrWhiteSpace(entry.Word)) return;
            string key = TextHelper.NormalizeWord(entry.Word);
            _entries[key] = entry;
            Save();
            Logger.Info($"РР·РѕТі СЃР°Т›Р»Р°РЅРґРё: '{key}'");
        }

        /// <summary>
        /// Adds or updates multiple explanation entries and saves once.
        /// Returns count of applied entries.
        /// </summary>
        public int AddOrUpdateBatch(IEnumerable<ExplanationEntry> entries)
        {
            if (entries == null) return 0;

            int applied = 0;
            foreach (var entry in entries)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.Word)) continue;

                string key = TextHelper.NormalizeWord(entry.Word);
                if (string.IsNullOrWhiteSpace(key)) continue;

                _entries[key] = entry;
                applied++;
            }

            if (applied > 0)
            {
                Save();
                Logger.Info($"РР·РѕТі РїР°РєРµС‚ СЃР°Т›Р»Р°РЅРґРё: {applied} С‚Р°");
            }

            return applied;
        }

        /// <summary>
        /// Removes the explanation entry for a word.
        /// </summary>
        public void Remove(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return;
            string key = TextHelper.NormalizeWord(word);
            if (_entries.Remove(key))
            {
                Save();
                Logger.Info($"РР·РѕТі СћС‡РёСЂРёР»РґРё: '{key}'");
            }
        }

        /// <summary>
        /// Loads the explanation database from the JSON file.
        /// </summary>
        public void Load()
        {
            try
            {
                _entries.Clear();

                if (!File.Exists(_dataPath))
                {
                    Logger.Warn($"РР·РѕТіР»Р°СЂ С„Р°Р№Р»Рё С‚РѕРїРёР»РјР°РґРё: {_dataPath}");
                    return;
                }

                string json = File.ReadAllText(_dataPath, Encoding.UTF8);
                var entries = ParseExplanationsJson(json);

                foreach (var e in entries)
                {
                    if (!string.IsNullOrWhiteSpace(e.Word))
                    {
                        string key = TextHelper.NormalizeWord(e.Word);
                        _entries[key] = e;
                    }
                }

                Logger.Info($"РР·РѕТіР»Р°СЂ СЋРєР»Р°РЅРґРё: {_entries.Count} С‚Р° СЃСћР·");
            }
            catch (Exception ex)
            {
                Logger.Error("РР·РѕТіР»Р°СЂ Р±Р°Р·Р°СЃРёРЅРё СЋРєР»Р°С€РґР° С…Р°С‚РѕР»РёРє", ex);
            }
        }

        /// <summary>
        /// Persists all explanation entries to the JSON file.
        /// </summary>
        public void Save()
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("[");
                bool first = true;
                foreach (var kvp in _entries)
                {
                    if (!first) sb.AppendLine(",");
                    first = false;
                    sb.AppendLine("  {");
                    sb.AppendFormat("    \"Word\": \"{0}\"", EscapeJson(kvp.Value.Word ?? kvp.Key));
                    if (!string.IsNullOrWhiteSpace(kvp.Value.Definition))
                        sb.AppendFormat(",\n    \"Definition\": \"{0}\"", EscapeJson(kvp.Value.Definition));
                    if (!string.IsNullOrWhiteSpace(kvp.Value.SpellingRule))
                        sb.AppendFormat(",\n    \"SpellingRule\": \"{0}\"", EscapeJson(kvp.Value.SpellingRule));
                    if (!string.IsNullOrWhiteSpace(kvp.Value.GrammarNote))
                        sb.AppendFormat(",\n    \"GrammarNote\": \"{0}\"", EscapeJson(kvp.Value.GrammarNote));
                    if (kvp.Value.Examples != null && kvp.Value.Examples.Length > 0)
                    {
                        sb.Append(",\n    \"Examples\": [");
                        for (int i = 0; i < kvp.Value.Examples.Length; i++)
                        {
                            if (i > 0) sb.Append(", ");
                            sb.AppendFormat("\"{0}\"", EscapeJson(kvp.Value.Examples[i]));
                        }
                        sb.Append("]");
                    }
                    sb.AppendLine();
                    sb.Append("  }");
                }
                sb.AppendLine();
                sb.AppendLine("]");

                string dir = Path.GetDirectoryName(_dataPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                File.WriteAllText(_dataPath, sb.ToString(), Encoding.UTF8);
                Logger.Info($"РР·РѕТіР»Р°СЂ СЃР°Т›Р»Р°РЅРґРё: {_entries.Count} С‚Р° СЃСћР·");
            }
            catch (Exception ex)
            {
                Logger.Error("РР·РѕТіР»Р°СЂРЅРё СЃР°Т›Р»Р°С€РґР° С…Р°С‚РѕР»РёРє", ex);
            }
        }

        private static ExplanationEntry FindExactOrStem(
            Dictionary<string, ExplanationEntry> entries,
            string normalizedWord)
        {
            if (entries == null || string.IsNullOrWhiteSpace(normalizedWord)) return null;

            ExplanationEntry entry;
            if (entries.TryGetValue(normalizedWord, out entry))
                return entry;

            string stemmed = RemoveCommonSuffixes(normalizedWord);
            if (!string.IsNullOrEmpty(stemmed) && stemmed != normalizedWord &&
                entries.TryGetValue(stemmed, out entry))
            {
                return entry;
            }

            return null;
        }

        private void EnsureBundledEntriesLoaded()
        {
            if (_bundledEntriesLoaded) return;

            lock (_bundledLoadLock)
            {
                if (_bundledEntriesLoaded) return;

                try
                {
                    _ensureBundledDictionary?.Invoke();
                    if (string.IsNullOrWhiteSpace(_bundledDictionaryPath) ||
                        !File.Exists(_bundledDictionaryPath))
                    {
                        return;
                    }

                    var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                    var payload = serializer.Deserialize<BundledDictionaryPayload>(
                        File.ReadAllText(_bundledDictionaryPath, Encoding.UTF8));

                    var entries = new Dictionary<string, ExplanationEntry>(StringComparer.OrdinalIgnoreCase);
                    if (payload?.Entries != null)
                    {
                        foreach (var source in payload.Entries)
                        {
                            if (source == null || string.IsNullOrWhiteSpace(source.Word) ||
                                !HasExplanationContent(source))
                            {
                                continue;
                            }

                            string key = TextHelper.NormalizeWord(source.Word);
                            if (string.IsNullOrWhiteSpace(key)) continue;

                            entries[key] = new ExplanationEntry
                            {
                                Word = source.Word,
                                Definition = source.Definition,
                                SpellingRule = source.SpellingRule,
                                GrammarNote = source.GrammarNote,
                                Examples = source.Examples != null ? source.Examples.ToArray() : null
                            };
                        }
                    }

                    _bundledEntries = entries;
                    Logger.Info($"Bundled dictionary metadata loaded lazily: {_bundledEntries.Count} entries");
                }
                catch (Exception ex)
                {
                    Logger.Error("Failed to load bundled dictionary metadata", ex);
                }
                finally
                {
                    // A failed optional fallback must not be retried for every
                    // lookup; the local explanation store remains available.
                    _bundledEntriesLoaded = true;
                }
            }
        }

        private static bool HasExplanationContent(ExplanationEntry entry)
        {
            return !string.IsNullOrWhiteSpace(entry.Definition) ||
                   !string.IsNullOrWhiteSpace(entry.SpellingRule) ||
                   !string.IsNullOrWhiteSpace(entry.GrammarNote) ||
                   (entry.Examples != null && entry.Examples.Any(example =>
                       !string.IsNullOrWhiteSpace(example)));
        }

        private static string EscapeJson(string s)
        {
            if (s == null) return "";
            return s.Replace("\\", "\\\\")
                    .Replace("\"", "\\\"")
                    .Replace("\n", "\\n")
                    .Replace("\r", "\\r")
                    .Replace("\t", "\\t");
        }

        // =====================================================================
        //  BASIC UZBEK STEMMING
        // =====================================================================

        /// <summary>
        /// Removes common Uzbek suffixes to find the root word.
        /// Not linguistically perfect, but good enough for dictionary lookup.
        /// </summary>
        private static string RemoveCommonSuffixes(string word)
        {
            if (string.IsNullOrEmpty(word) || word.Length < 4) return word;

            // Ordered from longest to shortest for greedy matching
            string[] suffixes = new[]
            {
                // Verb suffixes
                "Р»Р°СЂРёРЅРё", "Р»Р°СЂРёРіР°", "Р»Р°СЂРёРЅРё", "Р»Р°СЂРґР°РЅ",
                "РјРѕТ›РґР°", "Р№Р°РїС‚Рё", "РіР°РЅРґР°", "РјР°РіР°РЅ",
                "Р»Р°СЂРё", "РЅРёРЅРі", "РґР°РіРё", "РґР°Р»Рё",
                "Р»РёРіРё", "Р»Р°СЂРё", "СѓС‡СѓРЅ",
                "Р»Р°СЂ", "РЅРёРЅРі", "РґР°РЅ", "РіР°С‡Р°", "Р±РёР»Р°РЅ",
                "РіР°РЅ", "РЅРёР№", "РІРёР№", "РёР№",
                "С‡Рё", "Р»Рё", "РЅРё", "РґР°", "РіР°",
                "СЃРё", "РёРј", "РёРЅРі",
            };

            foreach (var suffix in suffixes)
            {
                if (word.Length > suffix.Length + 2 && word.EndsWith(suffix))
                {
                    return word.Substring(0, word.Length - suffix.Length);
                }
            }

            return word;
        }

        // =====================================================================
        //  JSON PARSER (no external dependencies)
        // =====================================================================

        /// <summary>
        /// Parses the explanations JSON array. Handles the structure:
        /// [{ "Word": "...", "Definition": "...", "SpellingRule": "...", 
        ///    "GrammarNote": "...", "Examples": ["...", "..."] }]
        /// </summary>
        private static List<ExplanationEntry> ParseExplanationsJson(string json)
        {
            var result = new List<ExplanationEntry>();
            if (string.IsNullOrWhiteSpace(json)) return result;

            json = json.Trim();
            if (!json.StartsWith("[")) return result;

            // Split into individual objects by matching { ... }
            // Using a brace-depth tracker to handle nested arrays
            var objects = ExtractJsonObjects(json);

            foreach (var obj in objects)
            {
                var entry = new ExplanationEntry
                {
                    Word = ExtractJsonString(obj, "Word"),
                    Definition = ExtractJsonString(obj, "Definition"),
                    SpellingRule = ExtractJsonString(obj, "SpellingRule"),
                    GrammarNote = ExtractJsonString(obj, "GrammarNote"),
                    Examples = ExtractJsonStringArray(obj, "Examples")
                };

                if (!string.IsNullOrWhiteSpace(entry.Word))
                    result.Add(entry);
            }

            return result;
        }

        /// <summary>
        /// Extracts top-level JSON objects from an array string.
        /// Correctly handles nested brackets/arrays.
        /// </summary>
        private static List<string> ExtractJsonObjects(string json)
        {
            var objects = new List<string>();
            int depth = 0;
            int start = -1;
            bool inString = false;
            bool escaped = false;

            for (int i = 0; i < json.Length; i++)
            {
                char c = json[i];

                if (escaped)
                {
                    escaped = false;
                    continue;
                }

                if (c == '\\')
                {
                    escaped = true;
                    continue;
                }

                if (c == '"')
                {
                    inString = !inString;
                    continue;
                }

                if (inString) continue;

                if (c == '{')
                {
                    if (depth == 0) start = i;
                    depth++;
                }
                else if (c == '}')
                {
                    depth--;
                    if (depth == 0 && start >= 0)
                    {
                        objects.Add(json.Substring(start, i - start + 1));
                        start = -1;
                    }
                }
            }

            return objects;
        }

        /// <summary>
        /// Extracts a string value for a given key from a JSON object string.
        /// </summary>
        private static string ExtractJsonString(string json, string key)
        {
            var pattern = new Regex(
                $"\"{Regex.Escape(key)}\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"",
                RegexOptions.IgnoreCase);
            var match = pattern.Match(json);
            if (match.Success)
                return UnescapeJson(match.Groups[1].Value);
            return null;
        }

        /// <summary>
        /// Extracts a string array for a given key from a JSON object string.
        /// </summary>
        private static string[] ExtractJsonStringArray(string json, string key)
        {
            // Find "key": [...]
            var pattern = new Regex(
                $"\"{Regex.Escape(key)}\"\\s*:\\s*\\[(.*?)\\]",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var match = pattern.Match(json);
            if (!match.Success) return new string[0];

            string inner = match.Groups[1].Value;
            var items = new List<string>();

            // Extract each quoted string
            var strPattern = new Regex("\"((?:[^\"\\\\]|\\\\.)*)\"");
            foreach (Match m in strPattern.Matches(inner))
            {
                items.Add(UnescapeJson(m.Groups[1].Value));
            }

            return items.ToArray();
        }

        private static string UnescapeJson(string s)
        {
            if (s == null) return "";
            return s.Replace("\\\"", "\"")
                    .Replace("\\\\", "\\")
                    .Replace("\\n", "\n")
                    .Replace("\\r", "\r")
                    .Replace("\\t", "\t");
        }
    }
}

