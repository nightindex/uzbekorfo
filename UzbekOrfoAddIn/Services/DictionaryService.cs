using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Manages the main (built-in) Uzbek dictionary and the user's personal custom dictionary.
    /// Provides fast lookup via HashSet, supports import/export, and persists user changes.
    /// 
    /// Architecture: The main dictionary is loaded as-is from the .dic file.
    /// When a Latin word is added or imported, BOTH the original Latin form AND
    /// its Cyrillic transliteration are stored, so each word is directly findable
    /// in both scripts without on-the-fly transliteration.
    /// The Contains() method still falls back to transliteration for words that
    /// were stored before this dual-storage was introduced.
    /// </summary>
    public class DictionaryService : IDictionaryService
    {
        private readonly HashSet<string> _mainDictionary = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _userDictionary = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private readonly string _mainDictPath;
        private readonly string _userDictPath;

        /// <summary>
        /// Transliterator for Latin РІвЂ вЂќ Cyrillic conversion.
        /// Set after construction via <see cref="SetTransliterator"/>.
        /// </summary>
        private TransliterationService _transliterator;

        // ---- Performance: length-based search index for fast similarity lookups ----
        private Dictionary<int, List<string>> _mainByLength;
        private Dictionary<int, List<string>> _userByLength;

        // ---- Bigram inverted index: bigram → word list ----------------------------
        // Reduces FindSimilarWords candidate set from O(N/avgLen) to O(~1500)
        // at any dictionary size.
        // Main indexes: built once in Load(), never rebuilt (main dict is immutable).
        // User indexes: rebuilt asynchronously after every AddWord/RemoveWord.
        private Dictionary<string, List<string>> _mainBigramIndex;
        private Dictionary<string, List<string>> _userBigramIndex;

        // Lock guards _user/_mainWordsSortedCache and their dirty flags
        // against races between background WarmUpCaches and UI-thread callers.
        private readonly object _cacheLock = new object();
        private bool _userWordsCacheDirty = true;
        private List<string> _userWordsSortedCache = new List<string>();
        private bool _mainWordsCacheDirty = true;
        private List<string> _mainWordsSortedCache = new List<string>();
        private bool _allWordsCacheDirty = true;
        private List<string> _allWordsSortedCache = new List<string>();

        // ---- Editor pre-split cache: avoids re-transliterating ~200K words
        //      every time the dictionary editor is opened. Invalidated together
        //      with _allWordsCacheDirty. ------------------------------------------
        private EditorWordCache _editorCache;

        /// <summary>
        /// Pre-computed data for the dictionary editor form.
        /// Built once on a background thread; reused on subsequent opens.
        /// </summary>
        public sealed class EditorWordCache
        {
            public List<string> AllWords { get; set; }
            public List<string> CyrillicWords { get; set; }
            public List<string> LatinWords { get; set; }
            public Dictionary<string, string> LatinToCyrillicMap { get; set; }
        }

        public DictionaryService(string mainDictPath, string userDictPath)
        {
            _mainDictPath = mainDictPath;
            _userDictPath = userDictPath;
        }

        /// <summary>
        /// Injects the transliterator so that Contains/AddWord can be script-aware.
        /// Called after both DictionaryService and TransliterationService are initialized.
        /// </summary>
        public void SetTransliterator(TransliterationService transliterator)
        {
            _transliterator = transliterator;
            Logger.Info("Р СћРЎР‚Р В°Р Р…РЎРѓР В»Р С‘РЎвЂљР ВµРЎР‚Р В°РЎвЂљР С•РЎР‚ Р В»РЎС“РўвЂњР В°РЎвЂљР С–Р В° РЎС“Р В»Р В°Р Р…Р Т‘Р С‘ РІР‚вЂќ Р В»Р В°РЎвЂљР С‘Р Р…/Р С”Р С‘РЎР‚Р С‘Р В»Р В» Р В°Р Р†РЎвЂљР С•Р СР В°РЎвЂљР С‘Р С” Р В°Р В»Р СР В°РЎв‚¬Р С‘Р Р…Р В°Р Т‘Р С‘");
        }

        /// <inheritdoc/>
        public int TotalWordCount => _mainDictionary.Count + _userDictionary.Count;

        // =====================================================================
        //  LOAD / SAVE
        // =====================================================================

        /// <inheritdoc/>
        public void Load()
        {
            LoadMainDictionary();
            LoadUserDictionary();
            Logger.Info($"Р вЂєРЎС“РўвЂњР В°РЎвЂљ РЎР‹Р С”Р В»Р В°Р Р…Р Т‘Р С‘: Р В°РЎРѓР С•РЎРѓР С‘Р в„–={_mainDictionary.Count}, РЎв‚¬Р В°РЎвЂ¦РЎРѓР С‘Р в„–={_userDictionary.Count}");
            // Search indexes built on background thread via RebuildSearchIndexBackground()
        }

        private void LoadMainDictionary()
        {
            try
            {
                if (!File.Exists(_mainDictPath))
                {
                    Logger.Info($"Р С’РЎРѓР С•РЎРѓР С‘Р в„– Р В»РЎС“РўвЂњР В°РЎвЂљ РЎвЂљР С•Р С—Р С‘Р В»Р СР В°Р Т‘Р С‘: {_mainDictPath}");
                    return;
                }

                foreach (var line in File.ReadLines(_mainDictPath))
                {
                    var word = line.Trim().ToLowerInvariant();
                    if (!string.IsNullOrEmpty(word) && !word.StartsWith("#"))
                    {
                        _mainDictionary.Add(word);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Р С’РЎРѓР С•РЎРѓР С‘Р в„– Р В»РЎС“РўвЂњР В°РЎвЂљР Р…Р С‘ РЎР‹Р С”Р В»Р В°РЎв‚¬Р Т‘Р В° РЎвЂ¦Р В°РЎвЂљР С•", ex);
            }
        }

        private void LoadUserDictionary()
        {
            try
            {
                if (!File.Exists(_userDictPath)) return;

                foreach (var line in File.ReadLines(_userDictPath))
                {
                    var word = line.Trim().ToLowerInvariant();
                    if (!string.IsNullOrEmpty(word))
                    {
                        _userDictionary.Add(word);
                    }
                }
                _userWordsCacheDirty = true;
                _allWordsCacheDirty = true;
                _editorCache = null;
            }
            catch (Exception ex)
            {
                Logger.Error("Р РЃР В°РЎвЂ¦РЎРѓР С‘Р в„– Р В»РЎС“РўвЂњР В°РЎвЂљР Р…Р С‘ РЎР‹Р С”Р В»Р В°РЎв‚¬Р Т‘Р В° РЎвЂ¦Р В°РЎвЂљР С•", ex);
            }
        }

        /// <inheritdoc/>
        public void Save()
        {
            try
            {
                EnsureUserWordsCache();
                File.WriteAllLines(_userDictPath, _userWordsSortedCache);
                Logger.Info($"Р РЃР В°РЎвЂ¦РЎРѓР С‘Р в„– Р В»РЎС“РўвЂњР В°РЎвЂљ РЎРѓР В°РўвЂєР В»Р В°Р Р…Р Т‘Р С‘: {_userWordsSortedCache.Count} РЎРѓРЎС›Р В·");
            }
            catch (Exception ex)
            {
                Logger.Error("Р РЃР В°РЎвЂ¦РЎРѓР С‘Р в„– Р В»РЎС“РўвЂњР В°РЎвЂљР Р…Р С‘ РЎРѓР В°РўвЂєР В»Р В°РЎв‚¬Р Т‘Р В° РЎвЂ¦Р В°РЎвЂљР С•", ex);
            }
        }

        // =====================================================================
        //  LOOKUP
        // =====================================================================

        /// <inheritdoc/>
        /// <remarks>
        /// Script-aware lookup: if the word is Latin and not found directly,
        /// it is transliterated to Cyrillic and checked again (and vice versa).
        /// </remarks>
        public bool Contains(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return false;
            var normalized = TextHelper.NormalizeWord(word);

            // Direct lookup first
            if (_mainDictionary.Contains(normalized) || _userDictionary.Contains(normalized))
                return true;

            // If transliterator available, try the other script
            if (_transliterator != null)
            {
                string converted = ConvertToOtherScript(normalized);
                if (!string.IsNullOrEmpty(converted) && converted != normalized)
                {
                    if (_mainDictionary.Contains(converted) || _userDictionary.Contains(converted))
                        return true;
                }
            }

            return false;
        }

        // =====================================================================
        //  USER DICTIONARY MANAGEMENT
        // =====================================================================

        /// <inheritdoc/>
        /// <remarks>
        /// Stores the word in its canonical Cyrillic form. If the input is Latin,
        /// it is transliterated to Cyrillic before storing. This ensures one entry
        /// covers both scripts.
        /// Checks both main and user dictionaries first to prevent duplicates.
        /// </remarks>
        public AddWordResult AddWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return AddWordResult.Invalid;
            var normalized = TextHelper.NormalizeWord(word);
            if (string.IsNullOrEmpty(normalized)) return AddWordResult.Invalid;

            // Convert to Cyrillic canonical form if the word is Latin
            string canonical = ToCyrillicCanonical(normalized);
            // Cross-script form for fallback lookups (main dict is Latin-only)
            string otherScript = ConvertToOtherScript(normalized);

            // Check if already in main dictionary (either script + cross-script)
            if (_mainDictionary.Contains(canonical) || _mainDictionary.Contains(normalized) ||
                (!string.IsNullOrEmpty(otherScript) && otherScript != normalized && _mainDictionary.Contains(otherScript)))
            {
                Logger.Info($"Р РЋРЈР‡Р В· Р В°РЎРѓР С•РЎРѓР С‘Р в„– Р В»РЎС“РўвЂњР В°РЎвЂљР Т‘Р В° Р СР В°Р Р†Р В¶РЎС“Р Т‘: '{normalized}' ('{canonical}')");
                return AddWordResult.AlreadyInMainDictionary;
            }

            // Check if already in user dictionary (either script + cross-script)
            if (_userDictionary.Contains(canonical) || _userDictionary.Contains(normalized) ||
                (!string.IsNullOrEmpty(otherScript) && otherScript != normalized && _userDictionary.Contains(otherScript)))
            {
                Logger.Info($"Р РЋРЈР‡Р В· РЎв‚¬Р В°РЎвЂ¦РЎРѓР С‘Р в„– Р В»РЎС“РўвЂњР В°РЎвЂљР Т‘Р В° Р СР В°Р Р†Р В¶РЎС“Р Т‘: '{normalized}' ('{canonical}')");
                return AddWordResult.AlreadyInUserDictionary;
            }

            _userDictionary.Add(canonical);

            // Also store the original Latin form so both scripts are preserved
            if (!string.Equals(normalized, canonical, StringComparison.OrdinalIgnoreCase))
            {
                _userDictionary.Add(normalized);
            }

            _userWordsCacheDirty = true;
            _allWordsCacheDirty = true;
            _editorCache = null;
            RebuildUserIndexAsync();
            Save();
            Logger.Info($"Р вЂєРЎС“РўвЂњР В°РЎвЂљР С–Р В° РўвЂєРЎС›РЎв‚¬Р С‘Р В»Р Т‘Р С‘: '{normalized}' РІвЂ вЂ™ '{canonical}'");
            return AddWordResult.Added;
        }

        /// <inheritdoc/>
        public void RemoveWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return;
            var normalized = TextHelper.NormalizeWord(word);

            // Remove the word itself
            _userDictionary.Remove(normalized);

            // Also remove the canonical Cyrillic and the other-script form
            // so that both stored forms are cleaned up consistently
            string canonical = ToCyrillicCanonical(normalized);
            if (!string.Equals(canonical, normalized, StringComparison.OrdinalIgnoreCase))
                _userDictionary.Remove(canonical);

            string otherScript = ConvertToOtherScript(normalized);
            if (!string.IsNullOrEmpty(otherScript) &&
                !string.Equals(otherScript, normalized, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(otherScript, canonical, StringComparison.OrdinalIgnoreCase))
                _userDictionary.Remove(otherScript);

            _userWordsCacheDirty = true;
            _allWordsCacheDirty = true;
            _editorCache = null;
            RebuildUserIndexAsync();
            Save();
        }

        /// <inheritdoc/>
        public List<string> GetUserWords()
        {
            EnsureUserWordsCache();
            return new List<string>(_userWordsSortedCache);
        }

        private void EnsureUserWordsCache()
        {
            lock (_cacheLock)
            {
                if (!_userWordsCacheDirty) return;

                var list = new List<string>(_userDictionary);
                list.Sort(UzbekStringComparer.Instance);
                _userWordsSortedCache = list;
                _userWordsCacheDirty = false;
            }
        }

        /// <inheritdoc/>
        public List<string> GetMainWords()
        {
            EnsureMainWordsCache();
            return new List<string>(_mainWordsSortedCache);
        }

        private void EnsureMainWordsCache()
        {
            lock (_cacheLock)
            {
                if (!_mainWordsCacheDirty) return;

                var list = new List<string>(_mainDictionary);
                list.Sort(UzbekStringComparer.Instance);
                _mainWordsSortedCache = list;
                _mainWordsCacheDirty = false;
            }
        }

        /// <inheritdoc/>
        public List<string> GetAllWords()
        {
            EnsureAllWordsCache();
            return new List<string>(_allWordsSortedCache);
        }

        private void EnsureAllWordsCache()
        {
            lock (_cacheLock)
            {
                if (!_allWordsCacheDirty) return;

                var combined = new HashSet<string>(_mainDictionary, StringComparer.OrdinalIgnoreCase);
                foreach (var w in _userDictionary)
                    combined.Add(w);

                var list = new List<string>(combined);
                list.Sort(UzbekStringComparer.Instance);
                _allWordsSortedCache = list;
                _allWordsCacheDirty = false;
            }
        }

        /// <inheritdoc/>
        public bool IsMainDictionaryWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return false;
            var normalized = TextHelper.NormalizeWord(word);

            if (_mainDictionary.Contains(normalized))
                return true;

            // Cross-script fallback
            if (_transliterator != null)
            {
                string converted = ConvertToOtherScript(normalized);
                if (!string.IsNullOrEmpty(converted) && converted != normalized)
                {
                    if (_mainDictionary.Contains(converted))
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Returns the cached editor pre-split data, or null if it hasn't been
        /// built yet or was invalidated by a dictionary change.
        /// </summary>
        public EditorWordCache GetEditorCache()
        {
            lock (_cacheLock) { return _editorCache; }
        }

        /// <summary>
        /// Stores pre-split editor data so subsequent form opens are instant.
        /// Called by the form after the first background build completes.
        /// </summary>
        public void SetEditorCache(EditorWordCache cache)
        {
            lock (_cacheLock) { _editorCache = cache; }
        }

        /// <summary>
        /// Pre-builds both sorted caches on whichever thread this is called from.
        /// Sorts run in parallel (independent data) so at 2 M words the combined
        /// wall-clock time is roughly the cost of sorting the larger list alone.
        /// Call once at startup on a background thread.
        /// </summary>
        public void WarmUpCaches()
        {
            // Both Ensure*Cache methods are now synchronized via _cacheLock,
            // so they are safe to call from the background Task.Run thread
            // even if the UI thread calls them concurrently.
            EnsureUserWordsCache();
            EnsureMainWordsCache();
            EnsureAllWordsCache();
        }

        /// <summary>
        /// Pre-builds the editor word cache (transliteration + Cyrillic/Latin split)
        /// so that the dictionary editor opens instantly on first click.
        /// Call once on a background thread after WarmUpCaches().
        /// </summary>
        public void PreBuildEditorCache()
        {
            lock (_cacheLock)
            {
                if (_editorCache != null) return; // already built
            }

            var words = GetAllWords();
            var translit = _transliterator;
            var cyrList = new List<string>(words.Count);
            var latList = new List<string>(words.Count);
            var map = new Dictionary<string, string>(words.Count, StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < words.Count; i++)
            {
                string w = words[i];
                bool isCyr = false;
                foreach (char c in w)
                {
                    if (c >= '\u0400' && c <= '\u04FF') { isCyr = true; break; }
                }

                if (isCyr)
                {
                    cyrList.Add(w);
                    if (translit != null)
                    {
                        string lat = translit.ToLatin(w)?.ToLowerInvariant();
                        if (!string.IsNullOrEmpty(lat) && !map.ContainsKey(lat))
                        {
                            latList.Add(lat);
                            map[lat] = w;
                        }
                    }
                }
                else
                {
                    if (!map.ContainsKey(w))
                        latList.Add(w);
                }
            }

            latList.Sort(UzbekStringComparer.Instance);

            var cache = new EditorWordCache
            {
                AllWords = words,
                CyrillicWords = cyrList,
                LatinWords = latList,
                LatinToCyrillicMap = map
            };

            lock (_cacheLock) { _editorCache = cache; }
            Logger.Info($"Editor cache pre-built: {words.Count} total, {cyrList.Count} cyr, {latList.Count} lat");
        }

        // =====================================================================
        //  EXPORT (Migration)
        // =====================================================================

        public sealed class MigrationExportResult
        {
            public int WordCount { get; set; }
            public int DefinitionCount { get; set; }
        }

        private sealed class MigrationPayload
        {
            public string Schema { get; set; }
            public string ExportedAtUtc { get; set; }
            public int WordCount { get; set; }
            public int DefinitionCount { get; set; }
            public List<MigrationEntry> Entries { get; set; }
        }

        private sealed class MigrationEntry
        {
            public string Word { get; set; }
            public string Definition { get; set; }
            public string SpellingRule { get; set; }
            public string GrammarNote { get; set; }
            public string[] Examples { get; set; }
        }

        /// <summary>
        /// Exports user dictionary words only for migration.
        /// </summary>
        public int ExportForMigration(string filePath)
        {
            return ExportForMigration(filePath, null).WordCount;
        }

        /// <summary>
        /// Exports user dictionary words for migration.
        /// Supported formats: json, xlsx, xls.
        /// If explanation entries are provided, matching definitions are exported too.
        /// </summary>
        public MigrationExportResult ExportForMigration(string filePath, IEnumerable<ExplanationEntry> explanationEntries)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Экспорт файли йўли керак.", nameof(filePath));

            if (string.IsNullOrWhiteSpace(Path.GetExtension(filePath)))
                filePath += ".json";

            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            int exportedDefinitions;
            var exportEntries = BuildMigrationEntries(explanationEntries, out exportedDefinitions);

            switch (extension)
            {
                case ".json":
                    WriteMigrationJson(filePath, exportEntries, exportedDefinitions);
                    break;
                case ".xlsx":
                    WriteMigrationXlsx(filePath, exportEntries);
                    break;
                case ".xls":
                    WriteMigrationXls(filePath, exportEntries);
                    break;
                default:
                    throw new NotSupportedException("Экспорт формати қўллаб-қувватланмайди: " + extension);
            }

            Logger.Info(
                $"Экспорт: {Path.GetFileName(filePath)} | Сўз: {exportEntries.Count} | Изоҳ: {exportedDefinitions}");

            return new MigrationExportResult
            {
                WordCount = exportEntries.Count,
                DefinitionCount = exportedDefinitions
            };
        }

        private List<MigrationEntry> BuildMigrationEntries(
            IEnumerable<ExplanationEntry> explanationEntries,
            out int exportedDefinitions)
        {
            EnsureUserWordsCache();

            var explanationMap = BuildExplanationMap(explanationEntries);
            var exportEntries = new List<MigrationEntry>(_userWordsSortedCache.Count);
            exportedDefinitions = 0;

            for (int i = 0; i < _userWordsSortedCache.Count; i++)
            {
                string userWord = _userWordsSortedCache[i];
                var entry = new MigrationEntry { Word = userWord };

                var explanation = ResolveExplanationForWord(userWord, explanationMap);
                if (explanation != null && HasExplanationContent(explanation))
                {
                    entry.Definition = string.IsNullOrWhiteSpace(explanation.Definition)
                        ? null
                        : explanation.Definition.Trim();
                    entry.SpellingRule = string.IsNullOrWhiteSpace(explanation.SpellingRule)
                        ? null
                        : explanation.SpellingRule.Trim();
                    entry.GrammarNote = string.IsNullOrWhiteSpace(explanation.GrammarNote)
                        ? null
                        : explanation.GrammarNote.Trim();
                    entry.Examples = NormalizeExamples(explanation.Examples);
                    exportedDefinitions++;
                }

                exportEntries.Add(entry);
            }

            return exportEntries;
        }

        private void WriteMigrationJson(string filePath, List<MigrationEntry> exportEntries, int exportedDefinitions)
        {
            var payload = new MigrationPayload
            {
                Schema = "uzbekorfo-migration-v1",
                ExportedAtUtc = DateTime.UtcNow.ToString("o"),
                WordCount = exportEntries.Count,
                DefinitionCount = exportedDefinitions,
                Entries = exportEntries
            };

            var serializer = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            };

            string output = PrettyPrintJson(serializer.Serialize(payload));
            File.WriteAllText(filePath, output, new UTF8Encoding(false));
        }

        private void WriteMigrationXlsx(string filePath, List<MigrationEntry> entries)
        {
            XlsxExportHelper.WriteXlsxPackage(filePath, "Dictionary", BuildMigrationSheetXml(entries));
        }

        private void WriteMigrationXls(string filePath, List<MigrationEntry> entries)
        {
            string tempXlsx = Path.Combine(
                Path.GetTempPath(),
                "uzbekorfo_dictionary_" + Guid.NewGuid().ToString("N") + ".xlsx");

            try
            {
                WriteMigrationXlsx(tempXlsx, entries);
                XlsxExportHelper.ConvertXlsxToXls(tempXlsx, filePath);
            }
            finally
            {
                try
                {
                    if (File.Exists(tempXlsx))
                        File.Delete(tempXlsx);
                }
                catch { }
            }
        }

        private string BuildMigrationSheetXml(List<MigrationEntry> entries)
        {
            if (entries == null) entries = new List<MigrationEntry>();

            var sb = new StringBuilder(8192);
            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
            sb.AppendLine(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            sb.AppendLine(@"  <sheetViews><sheetView workbookViewId=""0""/></sheetViews>");
            sb.AppendLine(@"  <sheetFormatPr defaultRowHeight=""15""/>");
            sb.AppendLine(@"  <cols>");
            sb.AppendLine(@"    <col min=""1"" max=""1"" width=""6"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""2"" max=""2"" width=""24"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""3"" max=""3"" width=""44"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""4"" max=""4"" width=""32"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""5"" max=""5"" width=""32"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""6"" max=""6"" width=""52"" customWidth=""1""/>");
            sb.AppendLine(@"  </cols>");
            sb.AppendLine(@"  <sheetData>");

            XlsxExportHelper.AppendInlineRow(sb, 1, new[] { "№", "Сўз", "Маъноси", "Имло қоидаси", "Грамматика", "Мисоллар" });

            int row = 2;
            for (int i = 0; i < entries.Count; i++)
            {
                var entry = entries[i] ?? new MigrationEntry();
                string examples = entry.Examples != null && entry.Examples.Length > 0
                    ? string.Join("; ", entry.Examples)
                    : string.Empty;

                XlsxExportHelper.AppendInlineRow(sb, row++, new[]
                {
                    (i + 1).ToString(),
                    entry.Word ?? string.Empty,
                    entry.Definition ?? string.Empty,
                    entry.SpellingRule ?? string.Empty,
                    entry.GrammarNote ?? string.Empty,
                    examples
                });
            }

            sb.AppendLine(@"  </sheetData>");
            sb.AppendLine(@"</worksheet>");
            return sb.ToString();
        }

        private Dictionary<string, ExplanationEntry> BuildExplanationMap(IEnumerable<ExplanationEntry> explanationEntries)
        {
            var map = new Dictionary<string, ExplanationEntry>(StringComparer.OrdinalIgnoreCase);
            if (explanationEntries == null) return map;

            foreach (var explanation in explanationEntries)
            {
                if (explanation == null || string.IsNullOrWhiteSpace(explanation.Word)) continue;
                if (!HasExplanationContent(explanation)) continue;

                string normalized = TextHelper.NormalizeWord(explanation.Word);
                if (string.IsNullOrWhiteSpace(normalized)) continue;

                string canonical = ToCyrillicCanonical(normalized);
                map[normalized] = explanation;
                map[canonical] = explanation;

                string opposite = ConvertToOtherScript(canonical);
                if (!string.IsNullOrWhiteSpace(opposite))
                    map[opposite] = explanation;
            }

            return map;
        }

        private ExplanationEntry ResolveExplanationForWord(
            string word,
            Dictionary<string, ExplanationEntry> explanationMap)
        {
            if (string.IsNullOrWhiteSpace(word) || explanationMap == null || explanationMap.Count == 0)
                return null;

            string normalized = TextHelper.NormalizeWord(word);
            if (string.IsNullOrWhiteSpace(normalized))
                return null;

            ExplanationEntry explanation;
            if (explanationMap.TryGetValue(normalized, out explanation))
                return explanation;

            string canonical = ToCyrillicCanonical(normalized);
            if (explanationMap.TryGetValue(canonical, out explanation))
                return explanation;

            string opposite = ConvertToOtherScript(normalized);
            if (!string.IsNullOrWhiteSpace(opposite) && explanationMap.TryGetValue(opposite, out explanation))
                return explanation;

            return null;
        }

        private static bool HasExplanationContent(ExplanationEntry explanation)
        {
            if (explanation == null) return false;

            return
                !string.IsNullOrWhiteSpace(explanation.Definition) ||
                !string.IsNullOrWhiteSpace(explanation.SpellingRule) ||
                !string.IsNullOrWhiteSpace(explanation.GrammarNote) ||
                (explanation.Examples != null && explanation.Examples.Any(v => !string.IsNullOrWhiteSpace(v)));
        }

        private static string PrettyPrintJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return string.Empty;

            var sb = new StringBuilder(json.Length + Math.Max(64, json.Length / 4));
            bool inQuotes = false;
            bool escaping = false;
            int indent = 0;

            for (int i = 0; i < json.Length; i++)
            {
                char ch = json[i];

                if (escaping)
                {
                    sb.Append(ch);
                    escaping = false;
                    continue;
                }

                if (ch == '\\' && inQuotes)
                {
                    sb.Append(ch);
                    escaping = true;
                    continue;
                }

                if (ch == '"')
                {
                    inQuotes = !inQuotes;
                    sb.Append(ch);
                    continue;
                }

                if (inQuotes)
                {
                    sb.Append(ch);
                    continue;
                }

                switch (ch)
                {
                    case '{':
                    case '[':
                        sb.Append(ch);
                        sb.AppendLine();
                        indent++;
                        sb.Append(new string(' ', indent * 2));
                        break;
                    case '}':
                    case ']':
                        sb.AppendLine();
                        indent = Math.Max(0, indent - 1);
                        sb.Append(new string(' ', indent * 2));
                        sb.Append(ch);
                        break;
                    case ',':
                        sb.Append(ch);
                        sb.AppendLine();
                        sb.Append(new string(' ', indent * 2));
                        break;
                    case ':':
                        sb.Append(": ");
                        break;
                    default:
                        if (!char.IsWhiteSpace(ch))
                            sb.Append(ch);
                        break;
                }
            }

            return sb.ToString();
        }

        // =====================================================================
        //  IMPORT
        // =====================================================================

        public sealed class ImportResult
        {
            public int ParsedWordCount { get; set; }
            public int AddedWordCount { get; set; }
            /// <summary>Words skipped because they already exist in the main dictionary.</summary>
            public int SkippedInMainCount { get; set; }
            /// <summary>Words skipped because they already exist in the user dictionary.</summary>
            public int SkippedInUserCount { get; set; }
            public int ProcessedFileCount { get; set; }
            public int FailedFileCount { get; set; }
            public int DefinitionCount => DefinitionEntries.Count;
            public List<ExplanationEntry> DefinitionEntries { get; } = new List<ExplanationEntry>();
        }

        private sealed class ImportedRow
        {
            public string Word { get; set; }
            public string Definition { get; set; }
            public string SpellingRule { get; set; }
            public string GrammarNote { get; set; }
            public string[] Examples { get; set; }
        }

        /// <inheritdoc/>
        public int ImportFromFile(string filePath)
        {
            return ImportFromFileDetailed(filePath).AddedWordCount;
        }

        /// <summary>
        /// Imports words from file and also extracts optional definitions/rules.
        /// Supported formats: txt, dic, csv, doc, docx, xls, xlsx, json.
        /// </summary>
        public ImportResult ImportFromFileDetailed(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Файл топилмади", filePath);

            var rows = ParseRowsFromFile(filePath);
            var result = new ImportResult();
            var definitionsByWord = new Dictionary<string, ExplanationEntry>(StringComparer.OrdinalIgnoreCase);

            ImportRows(rows, result, definitionsByWord);
            result.ProcessedFileCount = 1;

            if (result.AddedWordCount > 0)
            {
                _userWordsCacheDirty = true;
                _allWordsCacheDirty = true;
                _editorCache = null;
                RebuildUserIndexAsync();
                Save();
            }

            if (definitionsByWord.Count > 0)
                result.DefinitionEntries.AddRange(definitionsByWord.Values);

            Logger.Info(
                $"Импорт: {Path.GetFileName(filePath)} | Парс: {result.ParsedWordCount} | " +
                $"Янги сўз: {result.AddedWordCount} | Изоҳ: {result.DefinitionCount}");

            return result;
        }

        public ImportResult ImportFromFilesDetailed(IEnumerable<string> filePaths)
        {
            return ImportFromFilesDetailed(filePaths, null, null);
        }

        /// <summary>
        /// Imports words from multiple files with optional progress reporting and cancellation.
        /// </summary>
        /// <param name="filePaths">Files to import.</param>
        /// <param name="progressCallback">Called after each file: (processedCount, totalCount, currentFileName).</param>
        /// <param name="isCancelled">If non-null, checked before each file; returns early when true.</param>
        public ImportResult ImportFromFilesDetailed(
            IEnumerable<string> filePaths,
            Action<int, int, string> progressCallback,
            Func<bool> isCancelled)
        {
            var result = new ImportResult();
            if (filePaths == null) return result;

            var fileList = filePaths as IList<string> ?? new List<string>(filePaths);
            int totalFiles = fileList.Count;

            var definitionsByWord = new Dictionary<string, ExplanationEntry>(StringComparer.OrdinalIgnoreCase);

            for (int idx = 0; idx < fileList.Count; idx++)
            {
                if (isCancelled != null && isCancelled())
                    break;

                var path = fileList[idx];
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    result.FailedFileCount++;
                    continue;
                }

                try
                {
                    string fileName = Path.GetFileName(path);
                    if (progressCallback != null)
                        progressCallback(idx + 1, totalFiles, fileName);

                    var rows = ParseRowsFromFile(path) ?? new List<ImportedRow>();
                    ImportRows(rows, result, definitionsByWord);
                    result.ProcessedFileCount++;
                }
                catch (Exception ex)
                {
                    result.FailedFileCount++;
                    Logger.Error($"Импорт файлда хато: {path}", ex);
                }
            }

            if (result.AddedWordCount > 0)
            {
                _userWordsCacheDirty = true;
                _allWordsCacheDirty = true;
                _editorCache = null;
                RebuildUserIndexAsync();
                Save();
            }

            if (definitionsByWord.Count > 0)
                result.DefinitionEntries.AddRange(definitionsByWord.Values);

            Logger.Info(
                $"Импорт (папка): файл={result.ProcessedFileCount} | " +
                $"Парс: {result.ParsedWordCount} | Янги сўз: {result.AddedWordCount} | Изоҳ: {result.DefinitionCount} | " +
                $"Хато файл: {result.FailedFileCount}");

            return result;
        }

        private void ImportRows(List<ImportedRow> rows, ImportResult result,
            Dictionary<string, ExplanationEntry> definitionsByWord)
        {
            if (rows == null || result == null || definitionsByWord == null) return;

            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                if (row == null || string.IsNullOrWhiteSpace(row.Word))
                    continue;

                string normalized = TextHelper.NormalizeWord(row.Word);
                if (string.IsNullOrWhiteSpace(normalized) || TextHelper.ShouldSkipWord(normalized))
                    continue;

                result.ParsedWordCount++;

                string canonical = ToCyrillicCanonical(normalized);
                string otherScript = ConvertToOtherScript(normalized);
                bool existsInMain = _mainDictionary.Contains(canonical) || _mainDictionary.Contains(normalized);
                bool existsInUser = _userDictionary.Contains(canonical) || _userDictionary.Contains(normalized);

                // Cross-script fallback: the main dictionary stores Latin words,
                // so a Cyrillic import must also check the Latin equivalent.
                if (!existsInMain && !string.IsNullOrEmpty(otherScript) && otherScript != normalized)
                {
                    existsInMain = _mainDictionary.Contains(otherScript);
                }
                if (!existsInUser && !string.IsNullOrEmpty(otherScript) && otherScript != normalized)
                {
                    existsInUser = _userDictionary.Contains(otherScript);
                }

                if (!existsInMain && !existsInUser)
                {
                    _userDictionary.Add(canonical);

                    // Also store the original Latin form so both scripts are preserved
                    if (!string.Equals(normalized, canonical, StringComparison.OrdinalIgnoreCase))
                    {
                        _userDictionary.Add(normalized);
                    }

                    result.AddedWordCount++;
                }
                else if (existsInMain)
                {
                    result.SkippedInMainCount++;
                }
                else
                {
                    result.SkippedInUserCount++;
                }

                var explanation = BuildExplanationEntry(row, canonical);
                if (explanation != null)
                {
                    string key = TextHelper.NormalizeWord(explanation.Word);
                    if (!string.IsNullOrWhiteSpace(key))
                        definitionsByWord[key] = explanation;
                }
            }
        }

        private static ExplanationEntry BuildExplanationEntry(ImportedRow row, string canonicalWord)
        {
            if (row == null) return null;

            string definition = (row.Definition ?? string.Empty).Trim();
            string spellingRule = (row.SpellingRule ?? string.Empty).Trim();
            string grammarNote = (row.GrammarNote ?? string.Empty).Trim();
            string[] examples = NormalizeExamples(row.Examples);

            bool hasContent =
                !string.IsNullOrWhiteSpace(definition) ||
                !string.IsNullOrWhiteSpace(spellingRule) ||
                !string.IsNullOrWhiteSpace(grammarNote) ||
                (examples != null && examples.Length > 0);

            if (!hasContent) return null;

            return new ExplanationEntry
            {
                Word = canonicalWord,
                Definition = definition,
                SpellingRule = spellingRule,
                GrammarNote = grammarNote,
                Examples = examples
            };
        }

        private List<ImportedRow> ParseRowsFromFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            switch (ext)
            {
                case ".doc":
                case ".docx":
                    return ParseRowsFromText(ReadWordDocument(filePath));
                case ".xls":
                    return ParseRowsFromXls(filePath);
                case ".xlsx":
                    return ParseRowsFromXlsx(filePath);
                case ".json":
                    return ParseRowsFromJson(filePath);
                case ".csv":
                    return ParseRowsFromDelimitedFile(filePath);
                case ".txt":
                case ".dic":
                default:
                    return ParseRowsFromText(File.ReadAllText(filePath));
            }
        }

        private List<ImportedRow> ParseRowsFromText(string text)
        {
            var rows = new List<ImportedRow>();
            if (string.IsNullOrWhiteSpace(text)) return rows;

            var lines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = (lines[i] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                ImportedRow structured;
                if (TryParseStructuredLine(line, out structured))
                {
                    rows.Add(structured);
                    continue;
                }

                var tokens = TextHelper.Tokenize(line);
                for (int t = 0; t < tokens.Count; t++)
                {
                    if (string.IsNullOrWhiteSpace(tokens[t].Normalized)) continue;
                    rows.Add(new ImportedRow { Word = tokens[t].Normalized });
                }
            }

            return rows;
        }

        private bool TryParseStructuredLine(string line, out ImportedRow row)
        {
            row = null;
            if (string.IsNullOrWhiteSpace(line)) return false;

            if (line.Contains("|") || line.Contains("\t"))
            {
                char delimiter = line.Contains("|") ? '|' : '\t';
                var fields = line.Split(new[] { delimiter }, StringSplitOptions.None)
                    .Select(v => (v ?? string.Empty).Trim())
                    .ToArray();

                if (fields.Length >= 2 && IsLikelyWord(fields[0]))
                {
                    row = new ImportedRow
                    {
                        Word = fields[0],
                        Definition = fields.Length > 1 ? fields[1] : null,
                        SpellingRule = fields.Length > 2 ? fields[2] : null,
                        GrammarNote = fields.Length > 3 ? fields[3] : null,
                        Examples = fields.Length > 4 ? SplitExamples(fields[4]) : null
                    };
                    return true;
                }
            }

            string[] separators = { " — ", " - ", ": ", " : " };
            for (int i = 0; i < separators.Length; i++)
            {
                string sep = separators[i];
                int idx = line.IndexOf(sep, StringComparison.Ordinal);
                if (idx <= 0) continue;

                string left = line.Substring(0, idx).Trim();
                string right = line.Substring(idx + sep.Length).Trim();
                if (IsLikelyWord(left) && !string.IsNullOrWhiteSpace(right))
                {
                    row = new ImportedRow
                    {
                        Word = left,
                        Definition = right
                    };
                    return true;
                }
            }

            return false;
        }

        private List<ImportedRow> ParseRowsFromDelimitedFile(string filePath)
        {
            var lines = File.ReadAllLines(filePath);
            if (lines == null || lines.Length == 0) return new List<ImportedRow>();

            char delimiter = DetectDelimiter(lines);
            var table = new List<string[]>();
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (string.IsNullOrWhiteSpace(line)) continue;
                table.Add(ParseDelimitedLine(line, delimiter));
            }

            return ParseRowsFromTable(table);
        }

        private sealed class HeaderMap
        {
            public int WordIndex = 0;
            public int DefinitionIndex = 1;
            public int SpellingIndex = 2;
            public int GrammarIndex = 3;
            public int ExamplesIndex = 4;
        }

        private List<ImportedRow> ParseRowsFromTable(IList<string[]> table)
        {
            var rows = new List<ImportedRow>();
            if (table == null || table.Count == 0) return rows;

            int wordIdx = 0;
            int defIdx = 1;
            int spellIdx = 2;
            int gramIdx = 3;
            int exIdx = 4;
            int startRow = 0;

            HeaderMap map;
            if (TryMapHeader(table[0], out map))
            {
                wordIdx = map.WordIndex;
                defIdx = map.DefinitionIndex;
                spellIdx = map.SpellingIndex;
                gramIdx = map.GrammarIndex;
                exIdx = map.ExamplesIndex;
                startRow = 1;
            }

            for (int i = startRow; i < table.Count; i++)
            {
                var cells = table[i];
                if (cells == null || cells.Length == 0) continue;

                string word = GetCell(cells, wordIdx);
                if (string.IsNullOrWhiteSpace(word)) continue;

                rows.Add(new ImportedRow
                {
                    Word = word,
                    Definition = GetCell(cells, defIdx),
                    SpellingRule = GetCell(cells, spellIdx),
                    GrammarNote = GetCell(cells, gramIdx),
                    Examples = SplitExamples(GetCell(cells, exIdx))
                });
            }

            return rows;
        }

        private bool TryMapHeader(string[] headerCells, out HeaderMap map)
        {
            map = new HeaderMap();
            if (headerCells == null || headerCells.Length == 0) return false;

            int word = -1, definition = -1, spelling = -1, grammar = -1, examples = -1;

            for (int i = 0; i < headerCells.Length; i++)
            {
                string cell = (headerCells[i] ?? string.Empty).Trim().ToLowerInvariant();
                if (string.IsNullOrEmpty(cell)) continue;

                if (word < 0 && ContainsAny(cell, "word", "сўз", "суз", "so'z", "soz", "term"))
                    word = i;
                else if (definition < 0 && ContainsAny(cell, "definition", "meaning", "изоҳ", "изох", "таъриф", "tarif", "ta'rif"))
                    definition = i;
                else if (spelling < 0 && ContainsAny(cell, "spelling", "имло", "imlo", "rule", "қоида", "qoida"))
                    spelling = i;
                else if (grammar < 0 && ContainsAny(cell, "grammar", "grammatika", "грамматика", "note"))
                    grammar = i;
                else if (examples < 0 && ContainsAny(cell, "example", "examples", "мисол", "misol"))
                    examples = i;
            }

            if (word < 0) return false;

            map.WordIndex = word;
            map.DefinitionIndex = definition >= 0 ? definition : map.DefinitionIndex;
            map.SpellingIndex = spelling >= 0 ? spelling : map.SpellingIndex;
            map.GrammarIndex = grammar >= 0 ? grammar : map.GrammarIndex;
            map.ExamplesIndex = examples >= 0 ? examples : map.ExamplesIndex;
            return true;
        }

        private static bool ContainsAny(string value, params string[] keys)
        {
            if (string.IsNullOrEmpty(value) || keys == null) return false;
            for (int i = 0; i < keys.Length; i++)
            {
                string key = keys[i];
                if (!string.IsNullOrEmpty(key) && value.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        private static string GetCell(string[] cells, int index)
        {
            if (cells == null || index < 0 || index >= cells.Length) return null;
            return (cells[index] ?? string.Empty).Trim();
        }

        private static bool IsLikelyWord(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            string v = value.Trim();
            if (v.Length > 100) return false;
            return v.Count(char.IsWhiteSpace) <= 2;
        }

        private static char DetectDelimiter(IEnumerable<string> lines)
        {
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;

                int comma = line.Count(c => c == ',');
                int semicolon = line.Count(c => c == ';');
                int tab = line.Count(c => c == '\t');

                if (tab >= semicolon && tab >= comma && tab > 0) return '\t';
                if (semicolon >= comma && semicolon > 0) return ';';
                if (comma > 0) return ',';
            }
            return ',';
        }

        private static string[] ParseDelimitedLine(string line, char delimiter)
        {
            var values = new List<string>();
            if (line == null) return values.ToArray();

            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                    continue;
                }

                if (c == delimiter && !inQuotes)
                {
                    values.Add(sb.ToString().Trim());
                    sb.Clear();
                    continue;
                }

                sb.Append(c);
            }

            values.Add(sb.ToString().Trim());
            return values.ToArray();
        }

        private List<ImportedRow> ParseRowsFromJson(string filePath)
        {
            var rows = new List<ImportedRow>();
            string json = File.ReadAllText(filePath);
            if (string.IsNullOrWhiteSpace(json)) return rows;

            try
            {
                var serializer = new JavaScriptSerializer
                {
                    MaxJsonLength = int.MaxValue
                };

                var root = serializer.DeserializeObject(json);
                ExtractJsonRows(root, rows);
            }
            catch (Exception ex)
            {
                Logger.Error($"JSON импорт хатоси: {filePath}", ex);
            }

            return rows;
        }

        private void ExtractJsonRows(object node, List<ImportedRow> rows)
        {
            if (node == null || rows == null) return;

            var textNode = node as string;
            if (!string.IsNullOrWhiteSpace(textNode))
            {
                rows.Add(new ImportedRow { Word = textNode.Trim() });
                return;
            }

            var arrayNode = node as object[];
            if (arrayNode != null)
            {
                for (int i = 0; i < arrayNode.Length; i++)
                    ExtractJsonRows(arrayNode[i], rows);
                return;
            }

            var dict = node as Dictionary<string, object>;
            if (dict == null) return;

            ImportedRow row;
            if (TryParseJsonWordObject(dict, out row))
                rows.Add(row);

            object nested;
            if (TryGetJsonValue(dict, "words", out nested) ||
                TryGetJsonValue(dict, "entries", out nested) ||
                TryGetJsonValue(dict, "items", out nested) ||
                TryGetJsonValue(dict, "data", out nested) ||
                TryGetJsonValue(dict, "rows", out nested))
            {
                ExtractJsonRows(nested, rows);
            }
        }

        private bool TryParseJsonWordObject(Dictionary<string, object> obj, out ImportedRow row)
        {
            row = null;
            if (obj == null) return false;

            string word = GetJsonString(obj, "word", "term", "text", "name");
            if (string.IsNullOrWhiteSpace(word)) return false;

            row = new ImportedRow
            {
                Word = word,
                Definition = GetJsonString(obj, "definition", "meaning", "izoh", "ta'rif", "tarif"),
                SpellingRule = GetJsonString(obj, "spellingRule", "spelling", "imlo", "rule"),
                GrammarNote = GetJsonString(obj, "grammarNote", "grammar", "note"),
                Examples = GetJsonExamples(obj, "examples", "misollar")
            };

            return true;
        }

        private static bool TryGetJsonValue(Dictionary<string, object> obj, string key, out object value)
        {
            value = null;
            if (obj == null || string.IsNullOrWhiteSpace(key)) return false;

            foreach (var kv in obj)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase))
                {
                    value = kv.Value;
                    return true;
                }
            }

            return false;
        }

        private static string GetJsonString(Dictionary<string, object> obj, params string[] keys)
        {
            if (obj == null || keys == null) return null;

            for (int i = 0; i < keys.Length; i++)
            {
                object value;
                if (!TryGetJsonValue(obj, keys[i], out value) || value == null)
                    continue;

                var text = value as string;
                if (!string.IsNullOrWhiteSpace(text))
                    return text.Trim();

                if (!(value is object[]) && !(value is Dictionary<string, object>))
                    return value.ToString();
            }

            return null;
        }

        private static string[] GetJsonExamples(Dictionary<string, object> obj, params string[] keys)
        {
            if (obj == null || keys == null) return null;

            for (int i = 0; i < keys.Length; i++)
            {
                object value;
                if (!TryGetJsonValue(obj, keys[i], out value) || value == null)
                    continue;

                var arr = value as object[];
                if (arr != null)
                {
                    var items = arr
                        .Where(v => v != null)
                        .Select(v => v.ToString())
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .Select(v => v.Trim())
                        .ToArray();
                    return items.Length > 0 ? items : null;
                }

                var text = value as string;
                if (!string.IsNullOrWhiteSpace(text))
                    return SplitExamples(text);
            }

            return null;
        }

        private List<ImportedRow> ParseRowsFromXls(string filePath)
        {
            var table = new List<string[]>();

            object excelApp = null;
            object workbooks = null;
            object workbook = null;
            object worksheets = null;
            object worksheet = null;
            object usedRange = null;

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    Logger.Warn("XLS импорт: Excel COM топилмади.");
                    return new List<ImportedRow>();
                }

                excelApp = Activator.CreateInstance(excelType);
                dynamic app = excelApp;
                app.Visible = false;
                app.DisplayAlerts = false;

                workbooks = app.Workbooks;
                dynamic books = workbooks;
                workbook = books.Open(filePath, Type.Missing, true);
                dynamic wb = workbook;

                worksheets = wb.Worksheets;
                dynamic sheets = worksheets;
                worksheet = sheets.Item[1];
                dynamic sheet = worksheet;

                usedRange = sheet.UsedRange;
                dynamic range = usedRange;
                object value2 = range.Value2;

                var values = value2 as object[,];
                if (values != null)
                {
                    int rowStart = values.GetLowerBound(0);
                    int rowEnd = values.GetUpperBound(0);
                    int colStart = values.GetLowerBound(1);
                    int colEnd = values.GetUpperBound(1);

                    for (int r = rowStart; r <= rowEnd; r++)
                    {
                        var row = new string[colEnd - colStart + 1];
                        bool hasAny = false;
                        for (int c = colStart; c <= colEnd; c++)
                        {
                            object cell = values[r, c];
                            string text = cell != null ? cell.ToString().Trim() : string.Empty;
                            row[c - colStart] = text;
                            if (!string.IsNullOrWhiteSpace(text)) hasAny = true;
                        }

                        if (hasAny)
                            table.Add(row);
                    }
                }
                else if (value2 != null)
                {
                    string text = value2.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                        table.Add(new[] { text });
                }

            }
            catch (Exception ex)
            {
                Logger.Error($"XLS импорт хатоси: {filePath}", ex);
            }
            finally
            {
                try
                {
                    if (workbook != null)
                        ((dynamic)workbook).Close(false);
                }
                catch { }

                try
                {
                    if (excelApp != null)
                        ((dynamic)excelApp).Quit();
                }
                catch { }

                ReleaseComObject(usedRange);
                ReleaseComObject(worksheet);
                ReleaseComObject(worksheets);
                ReleaseComObject(workbook);
                ReleaseComObject(workbooks);
                ReleaseComObject(excelApp);

                // Force release of any lingering COM RCWs held by the dynamic
                // dispatch layer.  Without this, EXCEL.EXE can survive as a
                // zombie process until the next GC cycle — which may never come
                // soon enough inside Word's managed host.
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return ParseRowsFromTable(table);
        }

        private List<ImportedRow> ParseRowsFromXlsx(string filePath)
        {
            var table = new List<string[]>();

            try
            {
                using (var fs = File.OpenRead(filePath))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    var shared = ReadSharedStrings(zip);
                    string sheetPath = ResolveFirstWorksheetPath(zip);
                    if (string.IsNullOrWhiteSpace(sheetPath))
                        return new List<ImportedRow>();

                    var sheetEntry = zip.GetEntry(sheetPath);
                    if (sheetEntry == null)
                        return new List<ImportedRow>();

                    using (var stream = sheetEntry.Open())
                    {
                        var doc = XDocument.Load(stream);
                        var ns = doc.Root != null ? doc.Root.Name.Namespace : XNamespace.None;
                        var rows = doc.Descendants(ns + "row");

                        foreach (var row in rows)
                        {
                            var values = new Dictionary<int, string>();
                            int maxCol = -1;

                            foreach (var cell in row.Elements(ns + "c"))
                            {
                                string cellRef = (string)cell.Attribute("r");
                                int col = GetColumnIndex(cellRef);
                                if (col < 0) continue;

                                values[col] = ReadXlsxCellValue(cell, ns, shared);
                                if (col > maxCol) maxCol = col;
                            }

                            if (maxCol < 0) continue;
                            var rowValues = new string[maxCol + 1];
                            foreach (var kv in values)
                                rowValues[kv.Key] = kv.Value;
                            table.Add(rowValues);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"XLSX импорт хатоси: {filePath}", ex);
            }

            return ParseRowsFromTable(table);
        }

        private static List<string> ReadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return list;

            using (var stream = entry.Open())
            {
                var doc = XDocument.Load(stream);
                var ns = doc.Root != null ? doc.Root.Name.Namespace : XNamespace.None;
                foreach (var si in doc.Descendants(ns + "si"))
                {
                    string text = string.Concat(si.Descendants(ns + "t").Select(t => (string)t));
                    list.Add(text ?? string.Empty);
                }
            }

            return list;
        }

        private static string ResolveFirstWorksheetPath(ZipArchive zip)
        {
            var workbook = zip.GetEntry("xl/workbook.xml");
            var rels = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (workbook == null || rels == null) return null;

            XDocument workbookDoc;
            XDocument relsDoc;
            using (var stream = workbook.Open()) workbookDoc = XDocument.Load(stream);
            using (var stream = rels.Open()) relsDoc = XDocument.Load(stream);

            var wbNs = workbookDoc.Root != null ? workbookDoc.Root.Name.Namespace : XNamespace.None;
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace pkgNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            var firstSheet = workbookDoc.Descendants(wbNs + "sheet").FirstOrDefault();
            if (firstSheet == null) return null;

            string relId = (string)firstSheet.Attribute(relNs + "id");
            if (string.IsNullOrWhiteSpace(relId)) return null;

            var rel = relsDoc.Descendants(pkgNs + "Relationship")
                .FirstOrDefault(r => string.Equals((string)r.Attribute("Id"), relId, StringComparison.OrdinalIgnoreCase));
            if (rel == null) return null;

            string target = (string)rel.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target)) return null;

            target = target.Replace('\\', '/');
            if (target.StartsWith("/"))
                target = target.TrimStart('/');
            else if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                target = "xl/" + target;

            return target;
        }

        private static int GetColumnIndex(string cellRef)
        {
            if (string.IsNullOrWhiteSpace(cellRef)) return -1;

            int col = 0;
            for (int i = 0; i < cellRef.Length; i++)
            {
                char c = cellRef[i];
                if (!char.IsLetter(c)) break;
                col = (col * 26) + (char.ToUpperInvariant(c) - 'A' + 1);
            }

            return col > 0 ? col - 1 : -1;
        }

        private static string ReadXlsxCellValue(XElement cell, XNamespace ns, List<string> sharedStrings)
        {
            if (cell == null) return string.Empty;

            string type = (string)cell.Attribute("t");
            if (string.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase))
            {
                string inline = string.Concat(cell.Descendants(ns + "t").Select(t => (string)t));
                return (inline ?? string.Empty).Trim();
            }

            string raw = ((string)cell.Element(ns + "v") ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            int idx;
            if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(raw, out idx) &&
                idx >= 0 && idx < sharedStrings.Count)
            {
                return (sharedStrings[idx] ?? string.Empty).Trim();
            }

            return raw;
        }

        private static string[] SplitExamples(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;

            var values = raw
                .Split(new[] { '\r', '\n', ';', '|', '•' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => (s ?? string.Empty).Trim())
                .Where(s => s.Length > 0)
                .ToArray();

            return values.Length > 0 ? values : null;
        }

        private static string[] NormalizeExamples(string[] values)
        {
            if (values == null || values.Length == 0) return null;

            var normalized = values
                .Select(v => (v ?? string.Empty).Trim())
                .Where(v => v.Length > 0)
                .ToArray();

            return normalized.Length > 0 ? normalized : null;
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null) return;
            try
            {
                if (Marshal.IsComObject(comObject))
                    Marshal.FinalReleaseComObject(comObject);
            }
            catch { }
        }

        private string ReadWordDocument(string filePath)
        {
            Word.Document doc = null;
            try
            {
                var app = Globals.ThisAddIn.Application;
                doc = app.Documents.Open(filePath, ReadOnly: true, Visible: false);
                string text = doc.Content.Text;
                return text;
            }
            catch (Exception ex)
            {
                Logger.Error("Word hujjatni o'qishda xato", ex);
                return string.Empty;
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(false); } catch { }
                    try { Marshal.ReleaseComObject(doc); } catch { }
                }
            }
        }
        // =====================================================================
        //  SUGGESTIONS HELPER
        // =====================================================================

        /// <summary>
        /// Finds words in both dictionaries that are within the given edit distance.
        /// Uses length-bucketed index for fast narrowing (avoids scanning all 80K+ words).
        /// </summary>
        public List<(string Word, int Distance)> FindSimilarWords(string word, int maxDistance = 2, int maxResults = 10)
        {
            var normalized = TextHelper.NormalizeWord(word);
            if (string.IsNullOrEmpty(normalized)) return new List<(string, int)>();

            var candidates = new List<(string Word, int Distance)>();

            // First check phonetic variants (very fast, high quality)
            var phoneticVariants = TextHelper.GeneratePhoneticVariants(normalized);
            foreach (var variant in phoneticVariants)
            {
                if (_mainDictionary.Contains(variant) || _userDictionary.Contains(variant))
                {
                    candidates.Add((variant, 1));
                }
            }

            // Edit distance search: bigram pre-filter for words ≥ 4 chars (400× fewer
            // BoundedEditDistance calls at 2 M-word scale), length-bucket fallback for short words.
            SearchWithBigramFilter(_mainBigramIndex, _mainByLength, normalized, maxDistance, candidates);
            SearchWithBigramFilter(_userBigramIndex, _userByLength, normalized, maxDistance, candidates);

            // Deduplicate and sort by distance, then alphabetically
            return candidates
                .GroupBy(c => c.Word)
                .Select(g => g.OrderBy(c => c.Distance).First())
                .OrderBy(c => c.Distance)
                .ThenBy(c => c.Word)
                .Take(maxResults)
                .ToList();
        }

        private void SearchDictionaryIndexed(Dictionary<int, List<string>> byLength,
            string word, int maxDistance, List<(string Word, int Distance)> results)
        {
            if (byLength == null) return;

            int minLen = Math.Max(1, word.Length - maxDistance);
            int maxLen = word.Length + maxDistance;

            for (int len = minLen; len <= maxLen; len++)
            {
                if (!byLength.TryGetValue(len, out var bucket)) continue;

                foreach (var entry in bucket)
                {
                    int dist = TextHelper.BoundedEditDistance(word, entry, maxDistance);
                    if (dist > 0 && dist <= maxDistance)
                    {
                        results.Add((entry, dist));
                    }
                }
            }
        }

        /// <summary>
        /// Builds/rebuilds both the length-bucketed index and the bigram inverted
        /// index for both dictionaries.  Called after Load() and on mutations.
        /// </summary>
        /// <summary>
        /// Full initial index build — called ONCE from background thread.
        /// Main-dictionary indexes are immutable after this point.
        /// User-dictionary indexes are rebuilt asynchronously by
        /// <see cref="RebuildUserIndexAsync"/> after every mutation.
        /// </summary>
        public void RebuildSearchIndexBackground()
        {
            _mainByLength    = BuildLengthIndex(_mainDictionary);
            _userByLength    = BuildLengthIndex(_userDictionary);
            _mainBigramIndex = BuildBigramIndex(_mainDictionary);
            _userBigramIndex = BuildBigramIndex(_userDictionary);
            Logger.Info("Search indexes built on background thread");
        }

        /// <summary>
        /// Rebuilds only the user-dictionary search indexes on a background thread.
        /// A snapshot of <see cref="_userDictionary"/> is taken on the calling thread
        /// (the Word COM STA thread) before the task starts, preventing any
        /// concurrent-modification exception.
        /// Reference assignments to the index fields are atomic on 64-bit .NET, so
        /// <see cref="FindSimilarWords"/> at worst reads one stale user-index entry
        /// immediately after a mutation — far better than stalling the UI thread for
        /// a full rebuild.
        /// The main-dictionary indexes (<see cref="_mainByLength"/>,
        /// <see cref="_mainBigramIndex"/>) are never touched here — they are built
        /// once at startup and remain valid for the lifetime of the add-in.
        /// </summary>
        private void RebuildUserIndexAsync()
        {
            // Snapshot on the calling (UI) thread before any async boundary.
            var snapshot = new HashSet<string>(_userDictionary, StringComparer.OrdinalIgnoreCase);
            System.Threading.Tasks.Task.Run(() =>
            {
                var lenIdx = BuildLengthIndex(snapshot);
                var bigIdx = BuildBigramIndex(snapshot);
                // Atomic reference swap — no lock required on x64 .NET.
                _userByLength    = lenIdx;
                _userBigramIndex = bigIdx;
            });
        }

        /// <summary>
        /// Builds a bigram → word-list index for fast similarity pre-filtering.
        /// Each word contributes (word.Length - 1) bigram entries.
        /// Memory: ~80 MB at 2 M words (10 M entries × 8 B ref + ~1 K unique keys).
        /// </summary>
        private static Dictionary<string, List<string>> BuildBigramIndex(
            HashSet<string> dictionary)
        {
            var index = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            foreach (var word in dictionary)
            {
                if (word.Length < 2) continue;
                for (int i = 0; i < word.Length - 1; i++)
                {
                    // Use a two-char key without allocating a substring where possible.
                    string bigram = word.Substring(i, 2);
                    if (!index.TryGetValue(bigram, out var bucket))
                        index[bigram] = bucket = new List<string>();
                    bucket.Add(word);
                }
            }
            return index;
        }

        /// <summary>
        /// Searches a bigram inverted index for words within <paramref name="maxDistance"/>.
        /// Gate: a valid candidate MUST share at least <c>max(1, word.Length - 1 - maxDistance)</c>
        /// bigrams with the query.  For a 7-char word at maxDistance 2 this threshold is 4,
        /// reducing ~700 K length-bucket candidates to ~1 500 on a 2 M-word dictionary.
        /// Short words (threshold &lt; 1) fall back to the length-bucket scan.
        /// </summary>
        private void SearchWithBigramFilter(
            Dictionary<string, List<string>> bigramIndex,
            Dictionary<int, List<string>> byLength,
            string word, int maxDistance,
            List<(string Word, int Distance)> results)
        {
            int threshold = word.Length - 1 - maxDistance; // min shared bigrams required
            if (threshold < 1 || bigramIndex == null)
            {
                // Word too short for effective bigram gating; use length buckets (cheap).
                SearchDictionaryIndexed(byLength, word, maxDistance, results);
                return;
            }

            // Count how many of the query's bigrams each candidate shares.
            var hits = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < word.Length - 1; i++)
            {
                string bigram = word.Substring(i, 2);
                if (!bigramIndex.TryGetValue(bigram, out var bucket)) continue;
                foreach (var candidate in bucket)
                {
                    hits.TryGetValue(candidate, out int c);
                    hits[candidate] = c + 1;
                }
            }

            // Only compute BoundedEditDistance for candidates that pass the bigram gate.
            foreach (var kv in hits)
            {
                if (kv.Value < threshold) continue;
                int dist = TextHelper.BoundedEditDistance(word, kv.Key, maxDistance);
                if (dist > 0 && dist <= maxDistance)
                    results.Add((kv.Key, dist));
            }
        }

        private static Dictionary<int, List<string>> BuildLengthIndex(HashSet<string> dictionary)
        {
            var index = new Dictionary<int, List<string>>();
            foreach (var word in dictionary)
            {
                int len = word.Length;
                if (!index.TryGetValue(len, out var list))
                {
                    list = new List<string>();
                    index[len] = list;
                }
                list.Add(word);
            }
            return index;
        }

        // =====================================================================
        //  SCRIPT CONVERSION HELPERS
        // =====================================================================

        /// <summary>
        /// Converts a normalized word to canonical Cyrillic form.
        /// If already Cyrillic or no transliterator, returns as-is.
        /// </summary>
        private string ToCyrillicCanonical(string normalizedWord)
        {
            if (_transliterator == null) return normalizedWord;

            var script = _transliterator.DetectScript(normalizedWord);
            if (script == ScriptType.Latin)
            {
                var cyrillic = _transliterator.ToCyrillic(normalizedWord)?.ToLowerInvariant();
                return !string.IsNullOrEmpty(cyrillic) ? cyrillic : normalizedWord;
            }
            return normalizedWord;
        }

        /// <summary>
        /// Converts a word to the opposite script for cross-script lookup.
        /// Latin РІвЂ вЂ™ Cyrillic, Cyrillic РІвЂ вЂ™ Latin.
        /// </summary>
        private string ConvertToOtherScript(string normalizedWord)
        {
            if (_transliterator == null) return null;

            var script = _transliterator.DetectScript(normalizedWord);
            switch (script)
            {
                case ScriptType.Latin:
                    return _transliterator.ToCyrillic(normalizedWord)?.ToLowerInvariant();
                case ScriptType.Cyrillic:
                    return _transliterator.ToLatin(normalizedWord)?.ToLowerInvariant();
                default:
                    return null;
            }
        }

        /// <summary>
        /// Converts a Cyrillic suggestion to the target script if needed.
        /// Used to return suggestions in the same script the user typed.
        /// </summary>
        public string ConvertSuggestionToScript(string suggestion, ScriptType targetScript)
        {
            if (_transliterator == null || targetScript == ScriptType.Unknown || targetScript == ScriptType.Mixed)
                return suggestion;

            var suggestionScript = _transliterator.DetectScript(suggestion);

            if (suggestionScript == ScriptType.Cyrillic && targetScript == ScriptType.Latin)
                return _transliterator.ToLatin(suggestion)?.ToLowerInvariant() ?? suggestion;

            if (suggestionScript == ScriptType.Latin && targetScript == ScriptType.Cyrillic)
                return _transliterator.ToCyrillic(suggestion)?.ToLowerInvariant() ?? suggestion;

            return suggestion;
        }

        /// <summary>
        /// Returns a stable lowercase representation for local language-model keys.
        /// Latin input is canonicalized to Cyrillic when the transliterator is available,
        /// matching the dictionary's cross-script lookup behavior.
        /// </summary>
        public string CanonicalizeForModel(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return string.Empty;

            string normalized = TextHelper.NormalizeWord(word);
            if (string.IsNullOrWhiteSpace(normalized)) return string.Empty;

            string canonical = ToCyrillicCanonical(normalized);
            return string.IsNullOrWhiteSpace(canonical)
                ? normalized.ToLowerInvariant()
                : canonical.ToLowerInvariant();
        }

        /// <summary>
        /// Detects the script of a word (delegates to transliterator).
        /// Returns Unknown if no transliterator is set.
        /// </summary>
        public ScriptType DetectWordScript(string word)
        {
            if (_transliterator == null || string.IsNullOrEmpty(word))
                return ScriptType.Unknown;
            return _transliterator.DetectScript(word);
        }
    }
}
