using System;
using System.Globalization;
using System.IO;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Manages application settings and file paths.
    /// All paths point to %AppData%/UzbekOrfo/ folder.
    /// Settings are persisted as a simple key=value file.
    /// </summary>
    public class SettingsManager
    {
        private const int MinPredictionsAllowed = 1;
        private const int MaxPredictionsAllowed = 10;
        private const int MinPredictionLengthAllowed = 1;
        private const int MaxPredictionLengthAllowed = 15;
        private const int MinSpellingSuggestionsAllowed = 1;
        private const int MaxSpellingSuggestionsAllowed = 20;

        // =====================================================================
        //  PATHS
        // =====================================================================

        private static readonly string AppDataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "UzbekOrfo");

        /// <summary>Root data directory for all Uzbek Orfo files.</summary>
        public string DataDirectory => AppDataDir;

        /// <summary>Path to the main (built-in) dictionary file.</summary>
        public string MainDictionaryPath => Path.Combine(AppDataDir, "uzbek_main.dic");

        /// <summary>Path to the lazily used generated built-in metadata index.</summary>
        public string DictionaryMetadataPath => Path.Combine(AppDataDir, "uzbek_dictionary_metadata.json");

        /// <summary>Path to the user's personal custom dictionary.</summary>
        public string UserDictionaryPath => Path.Combine(AppDataDir, "user_custom.dic");

        /// <summary>Path to transliteration exceptions JSON file.</summary>
        public string ExceptionsPath => Path.Combine(AppDataDir, "translit_exceptions.json");

        /// <summary>Path to the prediction model JSON file.</summary>
        public string PredictionModelPath => Path.Combine(AppDataDir, "prediction_model.json");

        /// <summary>Path to the frequency seed JSON file used by suggestion ranking.</summary>
        public string FrequencySeedPath => Path.Combine(AppDataDir, "uzbek_freq_seed.json");

        /// <summary>Path to the locally learned language-model JSON file.</summary>
        public string LanguageModelPath => Path.Combine(AppDataDir, "language_model.json");

        /// <summary>Path to the explanations database JSON file.</summary>
        public string ExplanationsPath => Path.Combine(AppDataDir, "explanations.json");

        /// <summary>Path to the suffix definitions for morphological analysis.</summary>
        public string SuffixesPath => Path.Combine(AppDataDir, "uzbek_suffixes.json");

        /// <summary>Path to the grammar rules definitions.</summary>
        public string GrammarRulesPath => Path.Combine(AppDataDir, "grammar_rules.json");

        /// <summary>Path to the proper nouns list.</summary>
        public string ProperNounsPath => Path.Combine(AppDataDir, "proper_nouns.json");

        /// <summary>Path to the settings file.</summary>
        public string SettingsFilePath => Path.Combine(AppDataDir, "settings.cfg");

        // =====================================================================
        //  SETTINGS PROPERTIES
        // =====================================================================

        /// <summary>Whether auto-correct mode is enabled.</summary>
        public bool AutoCorrectEnabled { get; set; } = false;

        /// <summary>Whether grammar checking is enabled.</summary>
        public bool GrammarCheckEnabled { get; set; } = true;

        /// <summary>Whether predictive typing is enabled.</summary>
        public bool PredictionsEnabled { get; set; } = false;

        /// <summary>Maximum number of prediction suggestions to show.</summary>
        public int MaxPredictions { get; set; } = 3;

        /// <summary>Minimum word length to trigger predictions.</summary>
        public int MinPredictionLength { get; set; } = 2;

        /// <summary>Whether to auto-learn from typed text.</summary>
        public bool AutoLearn { get; set; } = true;

        /// <summary>Whether local frequency/context signals may influence suggestion ranking.</summary>
        public bool LanguageModelEnabled { get; set; } = true;

        /// <summary>Preferred script type: "Auto", "Cyrillic", "Latin".</summary>
        public string PreferredScript { get; set; } = "Auto";

        /// <summary>Prediction sensitivity (0.0 - 1.0).</summary>
        public double PredictionSensitivity { get; set; } = 0.5;

        /// <summary>Maximum suggestions shown in the Suggestions dialog.</summary>
        public int MaxSpellingSuggestions { get; set; } = 5;

        // =====================================================================
        //  INITIALIZATION
        // =====================================================================

        public SettingsManager()
        {
            EnsureDirectoryExists();
            Load();
            ValidateAndNormalize();
        }

        /// <summary>
        /// Ensures the AppData directory and initial files exist.
        /// </summary>
        public void EnsureDirectoryExists()
        {
            try
            {
                Directory.CreateDirectory(AppDataDir);

                // Create empty user dictionary if it doesn't exist
                if (!File.Exists(UserDictionaryPath))
                    File.WriteAllText(UserDictionaryPath, "");

                // Create empty exceptions file if it doesn't exist
                if (!File.Exists(ExceptionsPath))
                    File.WriteAllText(ExceptionsPath, "[]");
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error("Failed to create data directory", ex);
            }
        }

        // =====================================================================
        //  LOAD / SAVE
        // =====================================================================

        /// <summary>
        /// Loads settings from the settings file.
        /// Falls back to temp file if main file is missing (crash recovery).
        /// </summary>
        public void Load()
        {
            try
            {
                string path = SettingsFilePath;

                // If main file is missing but .tmp exists (crash during save), recover
                if (!File.Exists(path))
                {
                    string tempPath = path + ".tmp";
                    if (File.Exists(tempPath))
                    {
                        try { File.Move(tempPath, path); }
                        catch { }
                    }
                }

                if (!File.Exists(path)) return;

                foreach (var line in File.ReadAllLines(path))
                {
                    var parts = line.Split(new[] { '=' }, 2);
                    if (parts.Length != 2) continue;

                    var key = parts[0].Trim();
                    var value = parts[1].Trim();

                    switch (key)
                    {
                        case "AutoCorrectEnabled":
                            AutoCorrectEnabled = ParseBool(value, AutoCorrectEnabled);
                            break;
                        case "GrammarCheckEnabled":
                            GrammarCheckEnabled = ParseBool(value, GrammarCheckEnabled);
                            break;
                        case "PredictionsEnabled":
                            PredictionsEnabled = ParseBool(value, PredictionsEnabled);
                            break;
                        case "MaxPredictions":
                            MaxPredictions = ParseInt(value, MaxPredictions);
                            break;
                        case "MinPredictionLength":
                            MinPredictionLength = ParseInt(value, MinPredictionLength);
                            break;
                        case "AutoLearn":
                            AutoLearn = ParseBool(value, AutoLearn);
                            break;
                        case "LanguageModelEnabled":
                            LanguageModelEnabled = ParseBool(value, LanguageModelEnabled);
                            break;
                        case "PreferredScript":
                            PreferredScript = string.IsNullOrWhiteSpace(value) ? PreferredScript : value.Trim();
                            break;
                        case "PredictionSensitivity":
                            PredictionSensitivity = ParseDoubleInvariant(value, PredictionSensitivity);
                            break;
                        case "MaxSpellingSuggestions":
                            MaxSpellingSuggestions = ParseInt(value, MaxSpellingSuggestions);
                            break;
                    }
                }

                ValidateAndNormalize();
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error("Failed to load settings", ex);
            }
        }

        /// <summary>
        /// Persists current settings to disk.
        /// Uses write-to-temp-then-rename for atomic writes, preventing
        /// corruption if multiple Word instances save concurrently.
        /// </summary>
        public void Save()
        {
            try
            {
                ValidateAndNormalize();

                var lines = new[]
                {
                    $"AutoCorrectEnabled={AutoCorrectEnabled}",
                    $"GrammarCheckEnabled={GrammarCheckEnabled}",
                    $"PredictionsEnabled={PredictionsEnabled}",
                    $"MaxPredictions={MaxPredictions}",
                    $"MinPredictionLength={MinPredictionLength}",
                    $"AutoLearn={AutoLearn}",
                    $"LanguageModelEnabled={LanguageModelEnabled}",
                    $"PreferredScript={PreferredScript}",
                    $"PredictionSensitivity={PredictionSensitivity.ToString(System.Globalization.CultureInfo.InvariantCulture)}",
                    $"MaxSpellingSuggestions={MaxSpellingSuggestions}"
                };

                Helpers.AtomicFile.WriteAllLines(SettingsFilePath, lines);
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error("Failed to save settings", ex);
            }
        }

        private static bool ParseBool(string value, bool fallback)
        {
            return bool.TryParse(value, out bool parsed) ? parsed : fallback;
        }

        private static int ParseInt(string value, int fallback)
        {
            return int.TryParse(value, out int parsed) ? parsed : fallback;
        }

        private static double ParseDoubleInvariant(string value, double fallback)
        {
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : fallback;
        }

        private void ValidateAndNormalize()
        {
            if (!IsPreferredScript(PreferredScript))
                PreferredScript = "Auto";

            if (PredictionSensitivity < 0.0) PredictionSensitivity = 0.0;
            if (PredictionSensitivity > 1.0) PredictionSensitivity = 1.0;

            if (MaxPredictions < MinPredictionsAllowed) MaxPredictions = MinPredictionsAllowed;
            if (MaxPredictions > MaxPredictionsAllowed) MaxPredictions = MaxPredictionsAllowed;

            if (MinPredictionLength < MinPredictionLengthAllowed) MinPredictionLength = MinPredictionLengthAllowed;
            if (MinPredictionLength > MaxPredictionLengthAllowed) MinPredictionLength = MaxPredictionLengthAllowed;

            if (MaxSpellingSuggestions < MinSpellingSuggestionsAllowed)
                MaxSpellingSuggestions = MinSpellingSuggestionsAllowed;
            if (MaxSpellingSuggestions > MaxSpellingSuggestionsAllowed)
                MaxSpellingSuggestions = MaxSpellingSuggestionsAllowed;
        }

        private static bool IsPreferredScript(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            string normalized = value.Trim();
            return normalized.Equals("Auto", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("Cyrillic", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("Latin", StringComparison.OrdinalIgnoreCase);
        }

    }
}
