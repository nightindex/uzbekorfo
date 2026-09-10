using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Local, on-device language model:
    /// - Seed priors (bootstrapped frequency weights)
    /// - Learned unigram/bigram counts
    /// - Accepted correction history
    /// </summary>
    public sealed class LanguageModelService
    {
        private const string LocalSchemaVersion = "uzbekorfo-language-model-v1";
        private const string SeedSchemaVersion = "uzbekorfo-freq-v1";

        private const int MaxUnigramEntries = 12000;
        private const int MaxBigramEntries = 30000;
        private const int MaxCorrectionEntries = 20000;

        private readonly SettingsManager _settings;
        private readonly DictionaryService _dictionary;
        private readonly object _sync = new object();
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer
        {
            MaxJsonLength = int.MaxValue
        };

        private Dictionary<string, double> _seedPriors = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, int> _unigramCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, int> _bigramCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, int> _acceptedCorrectionCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        private int _dirtyEvents;

        private sealed class SeedPayload
        {
            public string Version { get; set; }
            public Dictionary<string, double> Priors { get; set; }
        }

        private sealed class LocalModelPayload
        {
            public string Version { get; set; }
            public Dictionary<string, int> UnigramCounts { get; set; }
            public Dictionary<string, int> BigramCounts { get; set; }
            public Dictionary<string, int> AcceptedCorrectionCounts { get; set; }
        }

        public LanguageModelService(SettingsManager settings, DictionaryService dictionary)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _dictionary = dictionary ?? throw new ArgumentNullException(nameof(dictionary));
        }

        public void Load()
        {
            LoadSeedPriors();
            LoadLocalModel();
        }

        public void Save()
        {
            lock (_sync)
            {
                SaveNoLock();
            }
        }

        public double GetUnigramPriorScore(string word)
        {
            string canonical = CanonicalizeToken(word);
            if (string.IsNullOrEmpty(canonical)) return 0.0;

            lock (_sync)
            {
                double seedScore = 0.30;
                if (_seedPriors.TryGetValue(canonical, out double seed))
                    seedScore = Clamp01(seed);

                int localCount = 0;
                _unigramCounts.TryGetValue(canonical, out localCount);
                double localScore = localCount <= 0 ? 0.0 : (double)localCount / (localCount + 8.0);

                return Clamp01((seedScore * 0.65) + (localScore * 0.35));
            }
        }

        public double GetContextScore(string previousWord, string candidateWord, string nextWord)
        {
            double left = GetBigramScore(previousWord, candidateWord);
            double right = GetBigramScore(candidateWord, nextWord);

            if (left <= 0.0 && right <= 0.0) return 0.0;
            if (left <= 0.0) return right;
            if (right <= 0.0) return left;
            return (left + right) / 2.0;
        }

        public double GetBigramScore(string leftWord, string rightWord)
        {
            string left = CanonicalizeToken(leftWord);
            string right = CanonicalizeToken(rightWord);
            if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right))
                return 0.0;

            lock (_sync)
            {
                int pairCount = 0;
                int leftCount = 0;
                _bigramCounts.TryGetValue(BuildBigramKey(left, right), out pairCount);
                _unigramCounts.TryGetValue(left, out leftCount);

                // Add-one smoothing and mild scaling to [0..1].
                double probability = (pairCount + 1.0) / (leftCount + 48.0);
                return Clamp01(probability * 6.0);
            }
        }

        public double GetAcceptedCorrectionBoost(string sourceWord, string candidateWord)
        {
            string source = CanonicalizeToken(sourceWord);
            string candidate = CanonicalizeToken(candidateWord);
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(candidate))
                return 0.0;

            lock (_sync)
            {
                int count = 0;
                _acceptedCorrectionCounts.TryGetValue(BuildCorrectionKey(source, candidate), out count);
                if (count <= 0) return 0.0;

                // Slowly saturating boost capped for safety.
                double boost = Math.Log(1 + count, 2) * 0.03;
                return Math.Min(0.20, boost);
            }
        }

        public void RecordAcceptedCorrection(string sourceWord, string replacementWord, SuggestionContext context)
        {
            string source = CanonicalizeToken(sourceWord);
            string replacement = CanonicalizeToken(replacementWord);
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(replacement))
                return;

            lock (_sync)
            {
                Increment(_acceptedCorrectionCounts, BuildCorrectionKey(source, replacement), 1);
                Increment(_unigramCounts, replacement, 2);

                string prev = CanonicalizeToken(context != null ? context.PreviousWord : null);
                string next = CanonicalizeToken(context != null ? context.NextWord : null);

                if (!string.IsNullOrEmpty(prev))
                    Increment(_bigramCounts, BuildBigramKey(prev, replacement), 2);
                if (!string.IsNullOrEmpty(next))
                    Increment(_bigramCounts, BuildBigramKey(replacement, next), 1);

                _dirtyEvents++;
                PruneUnsafeGrowthNoLock();
                if (_dirtyEvents >= 8)
                {
                    SaveNoLock();
                    _dirtyEvents = 0;
                }
            }
        }

        private void LoadSeedPriors()
        {
            try
            {
                string path = _settings.FrequencySeedPath;
                if (!File.Exists(path))
                {
                    GenerateFallbackSeedFile(path);
                }

                if (!File.Exists(path))
                    return;

                string json = File.ReadAllText(path);
                if (string.IsNullOrWhiteSpace(json))
                {
                    GenerateFallbackSeedFile(path);
                    json = File.Exists(path) ? File.ReadAllText(path) : string.Empty;
                }

                var payload = _serializer.Deserialize<SeedPayload>(json);
                var priors = payload != null ? payload.Priors : null;
                if (priors == null || priors.Count == 0)
                {
                    GenerateFallbackSeedFile(path);
                    payload = _serializer.Deserialize<SeedPayload>(File.ReadAllText(path));
                    priors = payload != null ? payload.Priors : null;
                }

                var normalized = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                if (priors != null)
                {
                    foreach (var kv in priors)
                    {
                        string key = CanonicalizeToken(kv.Key);
                        if (string.IsNullOrEmpty(key)) continue;
                        normalized[key] = Clamp01(kv.Value);
                    }
                }

                lock (_sync)
                {
                    _seedPriors = normalized;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("LanguageModel: seed priors load failed", ex);
                lock (_sync)
                {
                    _seedPriors = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                }
            }
        }

        private void LoadLocalModel()
        {
            try
            {
                string path = _settings.LanguageModelPath;
                if (!File.Exists(path))
                    return;

                string json = File.ReadAllText(path);
                if (string.IsNullOrWhiteSpace(json))
                    return;

                var payload = _serializer.Deserialize<LocalModelPayload>(json);
                if (payload == null)
                    return;

                lock (_sync)
                {
                    _unigramCounts = NormalizeIntDictionary(payload.UnigramCounts);
                    _bigramCounts = NormalizeIntDictionary(payload.BigramCounts);
                    _acceptedCorrectionCounts = NormalizeIntDictionary(payload.AcceptedCorrectionCounts);
                    PruneUnsafeGrowthNoLock();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("LanguageModel: local model load failed", ex);
            }
        }

        private void SaveNoLock()
        {
            try
            {
                string path = _settings.LanguageModelPath;
                string dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir))
                    Directory.CreateDirectory(dir);

                var payload = new LocalModelPayload
                {
                    Version = LocalSchemaVersion,
                    UnigramCounts = new Dictionary<string, int>(_unigramCounts, StringComparer.OrdinalIgnoreCase),
                    BigramCounts = new Dictionary<string, int>(_bigramCounts, StringComparer.OrdinalIgnoreCase),
                    AcceptedCorrectionCounts = new Dictionary<string, int>(_acceptedCorrectionCounts, StringComparer.OrdinalIgnoreCase)
                };

                string json = _serializer.Serialize(payload);
                File.WriteAllText(path, PrettyPrintJson(json));
            }
            catch (Exception ex)
            {
                Logger.Error("LanguageModel: local model save failed", ex);
            }
        }

        private void GenerateFallbackSeedFile(string path)
        {
            try
            {
                var priors = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                var words = _dictionary.GetMainWords();
                int take = Math.Min(5000, words.Count);

                if (take > 0)
                {
                    for (int i = 0; i < take; i++)
                    {
                        string key = CanonicalizeToken(words[i]);
                        if (string.IsNullOrEmpty(key) || priors.ContainsKey(key))
                            continue;

                        // Fallback prior (uniform + mild rank decay).
                        double decay = (double)i / Math.Max(1, take - 1);
                        priors[key] = Clamp01(0.35 + ((1.0 - decay) * 0.25));
                    }
                }

                if (priors.Count == 0)
                {
                    // Absolute fallback in case dictionary is unavailable.
                    priors["default"] = 0.35;
                }

                var payload = new SeedPayload
                {
                    Version = SeedSchemaVersion,
                    Priors = priors
                };

                string dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir))
                    Directory.CreateDirectory(dir);

                string json = _serializer.Serialize(payload);
                File.WriteAllText(path, PrettyPrintJson(json));
            }
            catch (Exception ex)
            {
                Logger.Error("LanguageModel: fallback seed generation failed", ex);
            }
        }

        private void PruneUnsafeGrowthNoLock()
        {
            PruneIntDictionary(_unigramCounts, MaxUnigramEntries);
            PruneIntDictionary(_bigramCounts, MaxBigramEntries);
            PruneIntDictionary(_acceptedCorrectionCounts, MaxCorrectionEntries);
        }

        private static void PruneIntDictionary(Dictionary<string, int> map, int maxEntries)
        {
            if (map == null || map.Count <= maxEntries) return;

            int removeCount = map.Count - maxEntries;
            if (removeCount <= 0) return;

            var keysToRemove = map
                .OrderBy(kv => kv.Value)
                .ThenBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                .Take(removeCount)
                .Select(kv => kv.Key)
                .ToList();

            for (int i = 0; i < keysToRemove.Count; i++)
            {
                map.Remove(keysToRemove[i]);
            }
        }

        private Dictionary<string, int> NormalizeIntDictionary(Dictionary<string, int> source)
        {
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (source == null) return result;

            foreach (var kv in source)
            {
                if (kv.Value <= 0) continue;
                string key = kv.Key;
                if (string.IsNullOrWhiteSpace(key)) continue;
                result[key.Trim()] = kv.Value;
            }
            return result;
        }

        private static void Increment(Dictionary<string, int> map, string key, int amount)
        {
            if (map == null || string.IsNullOrEmpty(key)) return;
            int existing = 0;
            map.TryGetValue(key, out existing);
            map[key] = existing + amount;
        }

        private string CanonicalizeToken(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return string.Empty;
            string canonical = _dictionary.CanonicalizeForModel(word);
            return canonical ?? string.Empty;
        }

        private static string BuildBigramKey(string left, string right)
        {
            return left + "\u001F" + right;
        }

        private static string BuildCorrectionKey(string source, string target)
        {
            return source + "\u001E" + target;
        }

        private static double Clamp01(double value)
        {
            if (value < 0.0) return 0.0;
            if (value > 1.0) return 1.0;
            return value;
        }

        private static string PrettyPrintJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return "{}";
            try
            {
                const string indent = "  ";
                int level = 0;
                bool inString = false;
                var result = new System.Text.StringBuilder(json.Length + 128);

                for (int i = 0; i < json.Length; i++)
                {
                    char ch = json[i];
                    if (ch == '"' && (i == 0 || json[i - 1] != '\\'))
                        inString = !inString;

                    if (!inString)
                    {
                        if (ch == '{' || ch == '[')
                        {
                            result.Append(ch);
                            result.AppendLine();
                            level++;
                            result.Append(new string(' ', level * indent.Length));
                            continue;
                        }
                        if (ch == '}' || ch == ']')
                        {
                            result.AppendLine();
                            level = Math.Max(0, level - 1);
                            result.Append(new string(' ', level * indent.Length));
                            result.Append(ch);
                            continue;
                        }
                        if (ch == ',')
                        {
                            result.Append(ch);
                            result.AppendLine();
                            result.Append(new string(' ', level * indent.Length));
                            continue;
                        }
                        if (ch == ':')
                        {
                            result.Append(": ");
                            continue;
                        }
                    }

                    if (!inString && char.IsWhiteSpace(ch))
                        continue;

                    result.Append(ch);
                }

                return result.ToString();
            }
            catch
            {
                return json;
            }
        }
    }
}
