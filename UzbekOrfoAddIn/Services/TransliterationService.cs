using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Full implementation of Uzbek Latin ↔ Cyrillic transliteration.
    /// Based on the official O'zbekiston Respublikasi transliteration standard (2021).
    ///
    /// Key complexities handled:
    ///   - Digraphs (Sh→Ш, Ch→Ч, Ng→Нг, etc.) processed before single letters
    ///   - Apostrophe variants for O' and G' (ʻ, ', ', ʼ)
    ///   - Case preservation (all-caps, title-case, lower-case)
    ///   - Exception words (proper names, brands, abbreviations)
    ///   - Mixed-script safety: only converts characters of the source script
    /// </summary>
    public class TransliterationService : ITransliterator
    {
        // =====================================================================
        //  MAPPING TABLES
        // =====================================================================

        #region Latin → Cyrillic mappings (digraphs FIRST, then single chars)

        /// <summary>
        /// Ordered list of Latin→Cyrillic rules.  Digraphs must come before their
        /// component letters so "Sh" matches before "S" + "h".
        /// Each entry: (latinPattern, cyrillicReplacement)
        /// </summary>
        private static readonly List<(string Latin, string Cyrillic)> LatinToCyrillicMap =
            new List<(string, string)>
            {
                // === Digraphs (case: UPPER-UPPER, Title, lower) ===
                ("SH", "Ш"),  ("Sh", "Ш"),  ("sh", "ш"),
                ("CH", "Ч"),  ("Ch", "Ч"),  ("ch", "ч"),
                ("NG", "НГ"), ("Ng", "Нг"), ("ng", "нг"),
                ("YO", "Ё"),  ("Yo", "Ё"),  ("yo", "ё"),
                ("YA", "Я"),  ("Ya", "Я"),  ("ya", "я"),
                ("YU", "Ю"),  ("Yu", "Ю"),  ("yu", "ю"),
                ("YE", "Е"),  ("Ye", "Е"),  ("ye", "е"),   // word-initial Ye → Е

                // === Apostrophe-based letters (multiple apostrophe forms) ===
                // O' / Oʻ / O' → Ў
                ("O\u02BB", "Ў"), ("O'", "Ў"), ("O\u2018", "Ў"), ("O\u2019", "Ў"), ("O\u02BC", "Ў"),
                ("o\u02BB", "ў"), ("o'", "ў"), ("o\u2018", "ў"), ("o\u2019", "ў"), ("o\u02BC", "ў"),
                // G' / Gʻ / G' → Ғ
                ("G\u02BB", "Ғ"), ("G'", "Ғ"), ("G\u2018", "Ғ"), ("G\u2019", "Ғ"), ("G\u02BC", "Ғ"),
                ("g\u02BB", "ғ"), ("g'", "ғ"), ("g\u2018", "ғ"), ("g\u2019", "ғ"), ("g\u02BC", "ғ"),

                // === Tutuq belgisi (ъ) — standalone apostrophe between vowels ===
                // Handled separately in code (see ApplyTutuqBelgisi)

                // === Single letters (uppercase then lowercase) ===
                ("A", "А"), ("a", "а"),
                ("B", "Б"), ("b", "б"),
                ("D", "Д"), ("d", "д"),
                ("E", "Е"), ("e", "е"),
                ("F", "Ф"), ("f", "ф"),
                ("G", "Г"), ("g", "г"),
                ("H", "Ҳ"), ("h", "ҳ"),
                ("I", "И"), ("i", "и"),
                ("J", "Ж"), ("j", "ж"),
                ("K", "К"), ("k", "к"),
                ("L", "Л"), ("l", "л"),
                ("M", "М"), ("m", "м"),
                ("N", "Н"), ("n", "н"),
                ("O", "О"), ("o", "о"),
                ("P", "П"), ("p", "п"),
                ("Q", "Қ"), ("q", "қ"),
                ("R", "Р"), ("r", "р"),
                ("S", "С"), ("s", "с"),
                ("T", "Т"), ("t", "т"),
                ("U", "У"), ("u", "у"),
                ("V", "В"), ("v", "в"),
                ("X", "Х"), ("x", "х"),
                ("Y", "Й"), ("y", "й"),
                ("Z", "З"), ("z", "з"),

                // TS → Ц  (rare, but standard)
                ("TS", "Ц"), ("Ts", "Ц"), ("ts", "ц"),
            };

        #endregion

        #region Cyrillic → Latin mappings

        /// <summary>
        /// Full Cyrillic→Latin map (uppercase+lowercase pairs), kept for reference / ToCyrillic reverse.
        /// </summary>
        private static readonly List<(string Cyrillic, string Latin)> CyrillicToLatinMap =
            new List<(string, string)>
            {
                // === Multi-character Cyrillic ===
                ("НГ", "NG"), ("Нг", "Ng"), ("нг", "ng"),

                // === Special Cyrillic letters → Latin digraphs ===
                ("Ш", "Sh"), ("ш", "sh"),
                ("Ч", "Ch"), ("ч", "ch"),
                ("Ё", "Yo"), ("ё", "yo"),
                ("Я", "Ya"), ("я", "ya"),
                ("Ю", "Yu"), ("ю", "yu"),
                ("Ц", "Ts"), ("ц", "ts"),

                // Uzbek-specific characters
                ("Ў", "O\u02BB"), ("ў", "o\u02BB"),   // Ў → Oʻ
                ("Ғ", "G\u02BB"), ("ғ", "g\u02BB"),   // Ғ → Gʻ
                ("Ҳ", "H"),  ("ҳ", "h"),               // Ҳ → H
                ("Қ", "Q"),  ("қ", "q"),               // Қ → Q
                ("Ъ", "\u02BB"), ("ъ", "\u02BB"),      // Ъ → ʻ (tutuq belgisi)
                ("Ь", ""),   ("ь", ""),                 // Ь → dropped (soft sign not in Uzbek Latin)
                ("Щ", "Sh"), ("щ", "sh"),               // Practical borrowed-word handling
                ("Ы", "I"),  ("ы", "i"),                // Practical borrowed-word handling

                // === Single Cyrillic letters ===
                ("А", "A"), ("а", "a"),
                ("Б", "B"), ("б", "b"),
                ("В", "V"), ("в", "v"),
                ("Г", "G"), ("г", "g"),
                ("Д", "D"), ("д", "d"),
                ("Е", "E"), ("е", "e"),     // mid-word Е = e  (word-initial handled in TransliterateWordToLatin)
                ("Ж", "J"), ("ж", "j"),
                ("З", "Z"), ("з", "z"),
                ("И", "I"), ("и", "i"),
                ("Й", "Y"), ("й", "y"),
                ("К", "K"), ("к", "k"),
                ("Л", "L"), ("л", "l"),
                ("М", "M"), ("м", "m"),
                ("Н", "N"), ("н", "n"),
                ("О", "O"), ("о", "o"),
                ("П", "P"), ("п", "p"),
                ("Р", "R"), ("р", "r"),
                ("С", "S"), ("с", "s"),
                ("Т", "T"), ("т", "t"),
                ("У", "U"), ("у", "u"),
                ("Ф", "F"), ("ф", "f"),
                ("Х", "X"), ("х", "x"),
                ("Э", "E"), ("э", "e"),
            };

        /// <summary>
        /// Lowercase-only Cyrillic→Latin map used during word-level transliteration.
        /// The caller normalises to lowercase first, then restores original case after.
        /// Ь/ь is mapped to "" (empty) — the soft sign is not present in Uzbek Latin.
        /// </summary>
        private static readonly List<(string Cyrillic, string Latin)> CyrillicToLatinMapLower =
            new List<(string, string)>
            {
                ("нг", "ng"),                        // digraph — must precede н
                ("ш", "sh"), ("ч", "ch"),
                ("ё", "yo"), ("я", "ya"), ("ю", "yu"), ("ц", "ts"),
                ("ў", "o\u02BB"),                    // oʻ
                ("ғ", "g\u02BB"),                    // gʻ
                ("ҳ", "h"), ("қ", "q"),
                ("ъ", "\u02BB"),                     // tutuq belgisi
                ("ь", ""),                           // soft sign → dropped
                ("щ", "sh"),                         // borrowed-word handling
                ("ы", "i"),                          // borrowed-word handling
                ("а","a"),("б","b"),("в","v"),("г","g"),("д","d"),
                ("е","e"),("ж","j"),("з","z"),("и","i"),
                ("й","y"),("к","k"),("л","l"),("м","m"),("н","n"),
                ("о","o"),("п","p"),("р","r"),("с","s"),("т","t"),
                ("у","u"),("ф","f"),("х","x"),("э","e"),
            };

        /// <summary>Word case patterns detected during transliteration.</summary>
        private enum WordCase { Lower, Upper, Title, Mixed }

        #endregion

        // Vowels used for tutuq belgisi detection
        private static readonly HashSet<char> LatinVowels =
            new HashSet<char>("aeiouAEIOU".ToCharArray());

        private static readonly HashSet<char> ApostropheChars =
            new HashSet<char>(new[] { '\'', '\u02BB', '\u02BC', '\u2018', '\u2019' });

        private static readonly HashSet<char> CyrillicVowelsLower =
            new HashSet<char>(new[] { 'а', 'е', 'ё', 'и', 'о', 'у', 'ў', 'э', 'ю', 'я', 'ы' });

        // =====================================================================
        //  STATE
        // =====================================================================

        private readonly string _exceptionsPath;
        private List<TranslitException> _exceptions = new List<TranslitException>();

        // =====================================================================
        //  CONSTRUCTOR
        // =====================================================================

        public TransliterationService(string exceptionsPath)
        {
            _exceptionsPath = exceptionsPath;
            LoadExceptions();

#if DEBUG
            RunTransliterationSelfCheck();
#endif
        }

        // =====================================================================
        //  ITransliterator — ToCyrillic
        // =====================================================================

        /// <summary>
        /// Converts Latin Uzbek text to Cyrillic. Handles exceptions, digraphs,
        /// apostrophe variants, and tutuq belgisi.
        /// </summary>
        public string ToCyrillic(string latinText)
        {
            if (string.IsNullOrEmpty(latinText)) return latinText;

            // 1. Apply exception replacements first
            string text = ApplyExceptions(latinText, toCyrillic: true);

            // 2. Process token by token (words vs non-words) to preserve formatting
            var sb = new StringBuilder(text.Length);
            int i = 0;

            while (i < text.Length)
            {
                bool matched = false;

                // Try digraphs and multi-char patterns first (longest match)
                foreach (var rule in LatinToCyrillicMap)
                {
                    if (i + rule.Latin.Length <= text.Length &&
                        text.Substring(i, rule.Latin.Length) == rule.Latin)
                    {
                        // Special: apostrophe after O/o or G/g might be tutuq belgisi
                        // If it's between two vowels and not O'/G', treat as ъ
                        sb.Append(rule.Cyrillic);
                        i += rule.Latin.Length;
                        matched = true;
                        break;
                    }
                }

                if (!matched)
                {
                    // Standalone apostrophe -> tutuq belgisi only in valid word-internal context.
                    // Example: mas'ul -> масъул, but quotes in 'word' stay as quotes.
                    if (ShouldConvertStandaloneApostropheToHardSign(text, i))
                    {
                        sb.Append('ъ');
                        i++;
                    }
                    else if (ApostropheChars.Contains(text[i]))
                    {
                        sb.Append(text[i]);
                        i++;
                    }
                    else
                    {
                        sb.Append(text[i]);
                        i++;
                    }
                }
            }

            return sb.ToString();
        }

        private static bool ShouldConvertStandaloneApostropheToHardSign(string text, int apostropheIndex)
        {
            if (string.IsNullOrEmpty(text) ||
                apostropheIndex <= 0 ||
                apostropheIndex >= text.Length - 1)
            {
                return false;
            }

            char mark = text[apostropheIndex];
            if (!ApostropheChars.Contains(mark))
                return false;

            char left = text[apostropheIndex - 1];
            char right = text[apostropheIndex + 1];

            // Only convert when apostrophe is inside a Latin word.
            // Prevents over-conversion in quotes/punctuation.
            if (!TextHelper.IsLatin(left) || !TextHelper.IsLatin(right))
                return false;

            // O'/G' -> Ў/Ғ is handled by explicit map entries and should not reach here.
            // If reached due to unusual spacing/input, keep apostrophe unchanged.
            char leftLower = char.ToLowerInvariant(left);
            if (leftLower == 'o' || leftLower == 'g')
                return false;

            return true;
        }

        // =====================================================================
        //  ITransliterator — ToLatin
        // =====================================================================

        /// <summary>
        /// Converts Cyrillic Uzbek text to Latin with correct case preservation.
        /// <list type="bullet">
        ///   <item>ALL-CAPS "ЮРИДА" → "YURIDA" (not "YuRIDA")</item>
        ///   <item>Title "Юрида" → "Yurida"</item>
        ///   <item>word-internal soft sign dropped: "компьютер" → "kompyuter"</item>
        ///   <item>Word-initial Е → "Ye", mid-word Е → "e"</item>
        /// </list>
        /// </summary>
        public string ToLatin(string cyrillicText)
        {
            if (string.IsNullOrEmpty(cyrillicText)) return cyrillicText;

            // 1. Apply exception replacements (whole-word, case-preserving)
            string text = ApplyExceptions(cyrillicText, toCyrillic: false);

            // 2. Process character by character, collecting letter-runs as words.
            //    Non-letter characters (spaces, punctuation, digits) pass through unchanged.
            var sb = new StringBuilder(text.Length * 2);
            int i = 0;

            while (i < text.Length)
            {
                if (char.IsLetter(text[i]))
                {
                    int start = i;
                    while (i < text.Length && char.IsLetter(text[i])) i++;
                    sb.Append(TransliterateWordToLatin(text.Substring(start, i - start)));
                }
                else
                {
                    sb.Append(text[i++]);
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Transliterates a single Cyrillic word to Latin with case preservation.
        /// Steps: detect case → lowercase → handle word-initial Е → map → restore case.
        /// </summary>
        private static string TransliterateWordToLatin(string word)
        {
            if (string.IsNullOrEmpty(word)) return word;

            // Lexical orthography exceptions (Uzbek normative forms)
            // сентябр/сентябрь -> sentabr, октябр/октябрь -> oktabr
            string loweredWord = word.ToLowerInvariant();
            if (loweredWord == "сентябр" || loweredWord == "сентябрь")
                return RestoreWordCase("sentabr", DetectWordCase(word));
            if (loweredWord == "октябр" || loweredWord == "октябрь")
                return RestoreWordCase("oktabr", DetectWordCase(word));

            WordCase caseType = DetectWordCase(word);

            // Normalise to lowercase so the map always hits lowercase entries
            string lower = word.ToLowerInvariant();

            // Apply longest-match map
            var buf = new StringBuilder(lower.Length * 2);
            int i = 0;
            while (i < lower.Length)
            {
                // Context rule for Е/e:
                // - word start or after vowel/ъ/ь -> "ye"
                // - after consonant -> "e"
                if (lower[i] == 'е')
                {
                    bool useYe = i == 0 ||
                                 CyrillicVowelsLower.Contains(lower[i - 1]) ||
                                 lower[i - 1] == 'ъ' || lower[i - 1] == 'ь';
                    buf.Append(useYe ? "ye" : "e");
                    i++;
                    continue;
                }

                // Context rule for Ц/ц (loanwords):
                // after vowel -> "ts", otherwise -> "s"
                if (lower[i] == 'ц')
                {
                    bool afterVowel = i > 0 && CyrillicVowelsLower.Contains(lower[i - 1]);
                    buf.Append(afterVowel ? "ts" : "s");
                    i++;
                    continue;
                }

                bool matched = false;
                foreach (var rule in CyrillicToLatinMapLower)
                {
                    int rLen = rule.Cyrillic.Length;
                    if (i + rLen <= lower.Length &&
                        lower.Substring(i, rLen) == rule.Cyrillic)
                    {
                        buf.Append(rule.Latin);
                        i += rLen;
                        matched = true;
                        break;
                    }
                }
                if (!matched) buf.Append(lower[i++]);
            }

            return RestoreWordCase(buf.ToString(), caseType);
        }

        private static WordCase DetectWordCase(string word)
        {
            bool allUpper = true;
            bool firstUpper = char.IsUpper(word[0]);
            bool restLower = true;

            for (int j = 0; j < word.Length; j++)
            {
                char c = word[j];
                if (!char.IsLetter(c)) continue;
                if (char.IsLower(c)) allUpper = false;
                if (j > 0 && char.IsUpper(c)) restLower = false;
            }

            if (allUpper) return WordCase.Upper;
            if (firstUpper && restLower) return WordCase.Title;
            if (!firstUpper) return WordCase.Lower;
            return WordCase.Mixed;
        }

        private static string RestoreWordCase(string latin, WordCase caseType)
        {
            if (string.IsNullOrEmpty(latin)) return latin;
            switch (caseType)
            {
                case WordCase.Upper: return latin.ToUpperInvariant();
                case WordCase.Title: return char.ToUpperInvariant(latin[0]) + latin.Substring(1);
                default:             return latin;   // Lower or Mixed → keep as-is (already lowercase)
            }
        }

        // =====================================================================
        //  ITransliterator — DetectScript
        // =====================================================================

        /// <summary>
        /// Detects the dominant script by counting Cyrillic vs Latin characters.
        /// </summary>
        public ScriptType DetectScript(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return ScriptType.Unknown;

            int cyrillic = 0, latin = 0;

            foreach (char c in text)
            {
                if (TextHelper.IsCyrillic(c)) cyrillic++;
                else if (TextHelper.IsLatin(c)) latin++;
            }

            int total = cyrillic + latin;
            if (total == 0) return ScriptType.Unknown;

            double cyrPercent = (double)cyrillic / total;
            double latPercent = (double)latin / total;

            if (cyrPercent > 0.8) return ScriptType.Cyrillic;
            if (latPercent > 0.8) return ScriptType.Latin;
            if (total < 3) return ScriptType.Unknown;

            return ScriptType.Mixed;
        }

        // =====================================================================
        //  ITransliterator — Exception Management
        // =====================================================================

        public List<TranslitException> GetExceptions() =>
            _exceptions.ToList(); // return copy

        public void AddException(TranslitException exception)
        {
            if (exception == null || string.IsNullOrWhiteSpace(exception.Original)) return;

            // Update if exists, add if new
            var existing = _exceptions
                .FirstOrDefault(e => e.Original.Equals(exception.Original, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                existing.Replacement = exception.Replacement;
                existing.Category = exception.Category;
                existing.Description = exception.Description;
                existing.Enabled = exception.Enabled;
                existing.Note = exception.Note;
            }
            else
            {
                _exceptions.Add(exception);
            }
        }

        public void RemoveException(string originalText)
        {
            _exceptions.RemoveAll(e =>
                e.Original.Equals(originalText, StringComparison.OrdinalIgnoreCase));
        }

        public void LoadExceptions()
        {
            try
            {
                if (File.Exists(_exceptionsPath))
                {
                    string json = File.ReadAllText(_exceptionsPath, Encoding.UTF8);
                    _exceptions = SimpleJsonParser.ParseExceptions(json);
                }

                // Seed defaults if file missing or empty
                if (_exceptions == null || _exceptions.Count == 0)
                {
                    _exceptions = GetDefaultExceptions();
                    SaveExceptions();
                    Logger.Info($"Бошланғич истиснолар юкланди: {_exceptions.Count} та");
                }
                else
                {
                    int merged = MergeMissingDefaultExceptions();
                    if (merged > 0)
                    {
                        SaveExceptions();
                        Logger.Info($"Янги стандарт истиснолар қўшилди: {merged} та");
                    }
                    Logger.Info($"Транслитерация истисноси юкланди: {_exceptions.Count} та");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Истисноларни юклашда хатолик", ex);
                _exceptions = new List<TranslitException>();
            }
        }

        public void SaveExceptions()
        {
            try
            {
                string json = SimpleJsonParser.SerializeExceptions(_exceptions);
                File.WriteAllText(_exceptionsPath, json, Encoding.UTF8);
                Logger.Info($"Истиснолар сақланди: {_exceptions.Count} та");
            }
            catch (Exception ex)
            {
                Logger.Error("Истисноларни сақлашда хатолик", ex);
            }
        }

        private int MergeMissingDefaultExceptions()
        {
            var defaults = GetDefaultExceptions();
            int added = 0;

            foreach (var def in defaults)
            {
                bool exists = _exceptions.Any(e =>
                    e.Original.Equals(def.Original, StringComparison.OrdinalIgnoreCase));
                if (!exists)
                {
                    _exceptions.Add(def);
                    added++;
                }
            }

            return added;
        }

        /// <summary>
        /// Returns a list of commonly-used default exceptions (brands, tech terms, etc.).
        /// Seeded when the exceptions file does not yet exist.
        /// </summary>
        private static List<TranslitException> GetDefaultExceptions()
        {
            return new List<TranslitException>
            {
                // \u0411\u0440\u0435\u043d\u0434\u043b\u0430\u0440
                new TranslitException("Microsoft",  "\u041c\u0430\u0439\u043a\u0440\u043e\u0441\u043e\u0444\u0442",   "Brand", "\u0413\u043b\u043e\u0431\u0430\u043b \u0442\u0435\u0445\u043d\u043e\u043b\u043e\u0433\u0438\u044f \u043a\u043e\u043c\u043f\u0430\u043d\u0438\u044f\u0441\u0438"),
                new TranslitException("Google",     "\u0413\u0443\u0433\u043b",         "Brand", "\u049a\u0438\u0434\u0438\u0440\u0443\u0432 \u0442\u0438\u0437\u0438\u043c\u0438"),
                new TranslitException("iPhone",     "iPhone",        "Brand", "Apple \u0441\u043c\u0430\u0440\u0442\u0444\u043e\u043d\u0438"),
                new TranslitException("Facebook",   "\u0424\u0435\u0439\u0441\u0431\u0443\u043a",      "Brand", "\u0418\u0436\u0442\u0438\u043c\u043e\u0438\u0439 \u0442\u0430\u0440\u043c\u043e\u049b"),
                new TranslitException("Linux",      "Linux",         "Brand", "\u041e\u043f\u0435\u0440\u0430\u0446\u0438\u043e\u043d \u0442\u0438\u0437\u0438\u043c"),
                new TranslitException("Samsung",    "\u0421\u0430\u043c\u0441\u0443\u043d\u0433",      "Brand", "\u042d\u043b\u0435\u043a\u0442\u0440\u043e\u043d\u0438\u043a\u0430 \u043a\u043e\u043c\u043f\u0430\u043d\u0438\u044f\u0441\u0438"),
                new TranslitException("Apple",      "Apple",         "Brand", "\u0422\u0435\u0445\u043d\u043e\u043b\u043e\u0433\u0438\u044f \u043a\u043e\u043c\u043f\u0430\u043d\u0438\u044f\u0441\u0438"),
                new TranslitException("Windows",    "Windows",       "Brand", "\u041e\u043f\u0435\u0440\u0430\u0446\u0438\u043e\u043d \u0442\u0438\u0437\u0438\u043c"),
                new TranslitException("Android",    "Android",       "Brand", "\u041c\u043e\u0431\u0438\u043b \u043e\u043f\u0435\u0440\u0430\u0446\u0438\u043e\u043d \u0442\u0438\u0437\u0438\u043c"),
                new TranslitException("Telegram",   "\u0422\u0435\u043b\u0435\u0433\u0440\u0430\u043c",     "Brand", "\u0425\u0430\u0431\u0430\u0440 \u0430\u043b\u043c\u0430\u0448\u0438\u0448 \u043f\u043b\u0430\u0442\u0444\u043e\u0440\u043c\u0430\u0441\u0438"),
                new TranslitException("YouTube",    "YouTube",       "Brand", "\u0412\u0438\u0434\u0435\u043e \u043f\u043b\u0430\u0442\u0444\u043e\u0440\u043c\u0430\u0441\u0438"),
                new TranslitException("Instagram",  "\u0418\u043d\u0441\u0442\u0430\u0433\u0440\u0430\u043c",    "Brand", "\u0418\u0436\u0442\u0438\u043c\u043e\u0438\u0439 \u0442\u0430\u0440\u043c\u043e\u049b"),
                new TranslitException("WhatsApp",   "WhatsApp",      "Brand", "\u041c\u0435\u0441\u0441\u0435\u043d\u0436\u0435\u0440 \u0438\u043b\u043e\u0432\u0430\u0441\u0438"),
                // \u0422\u0435\u0445\u043d\u0438\u043a \u0430\u0442\u0430\u043c\u0430\u043b\u0430\u0440
                new TranslitException("Python",     "Python",        "Technical", "\u0414\u0430\u0441\u0442\u0443\u0440\u043b\u0430\u0448 \u0442\u0438\u043b\u0438"),
                new TranslitException("Java",       "Java",          "Technical", "\u0414\u0430\u0441\u0442\u0443\u0440\u043b\u0430\u0448 \u0442\u0438\u043b\u0438"),
                new TranslitException("JavaScript", "JavaScript",    "Technical", "\u0414\u0430\u0441\u0442\u0443\u0440\u043b\u0430\u0448 \u0442\u0438\u043b\u0438"),
                new TranslitException("SQL",        "SQL",           "Technical", "\u041c\u0430\u044a\u043b\u0443\u043c\u043e\u0442\u043b\u0430\u0440 \u0431\u0430\u0437\u0430\u0441\u0438 \u0442\u0438\u043b\u0438"),
                new TranslitException("API",        "API",           "Technical", "\u0414\u0430\u0441\u0442\u0443\u0440\u0438\u0439 \u0438\u043d\u0442\u0435\u0440\u0444\u0435\u0439\u0441"),
                new TranslitException("JSON",       "JSON",          "Technical", "\u041c\u0430\u044a\u043b\u0443\u043c\u043e\u0442 \u0444\u043e\u0440\u043c\u0430\u0442\u0438"),
                new TranslitException("HTML",       "HTML",          "Technical", "\u0411\u0435\u043b\u0433\u0438\u043b\u0430\u0448 \u0442\u0438\u043b\u0438"),
                new TranslitException("CSS",        "CSS",           "Technical", "\u0421\u0442\u0438\u043b\u043b\u0430\u0440 \u0442\u0438\u043b\u0438"),
                new TranslitException("HTTP",       "HTTP",          "Technical", "\u0412\u0435\u0431 \u043f\u0440\u043e\u0442\u043e\u043a\u043e\u043b\u0438"),
                new TranslitException("URL",        "URL",           "Technical", "\u0412\u0435\u0431 \u043c\u0430\u043d\u0437\u0438\u043b \u0444\u043e\u0440\u043c\u0430\u0442\u0438"),
                new TranslitException("PDF",        "PDF",           "Technical", "\u04b2\u0443\u0436\u0436\u0430\u0442 \u0444\u043e\u0440\u043c\u0430\u0442\u0438"),
                new TranslitException("USB",        "USB",           "Technical", "\u0410\u043f\u043f\u0430\u0440\u0430\u0442 \u0438\u043d\u0442\u0435\u0440\u0444\u0435\u0439\u0441\u0438"),
                new TranslitException("WiFi",       "WiFi",          "Technical", "\u0421\u0438\u043c\u0441\u0438\u0437 \u0442\u0430\u0440\u043c\u043e\u049b"),
                // \u0413\u0435\u043e\u0433\u0440\u0430\u0444\u0438\u043a \u043d\u043e\u043c\u043b\u0430\u0440
                new TranslitException("Toshkent",   "\u0422\u043e\u0448\u043a\u0435\u043d\u0442",      "Geographic", "\u040e\u0437\u0431\u0435\u043a\u0438\u0441\u0442\u043e\u043d \u043f\u043e\u0439\u0442\u0430\u0445\u0442\u0438"),
                new TranslitException("Samarqand",  "\u0421\u0430\u043c\u0430\u0440\u049b\u0430\u043d\u0434",    "Geographic", "\u0422\u0430\u0440\u0438\u0445\u0438\u0439 \u0448\u0430\u04b3\u0430\u0440"),
                new TranslitException("Buxoro",     "\u0411\u0443\u0445\u043e\u0440\u043e",       "Geographic", "\u0422\u0430\u0440\u0438\u0445\u0438\u0439 \u0448\u0430\u04b3\u0430\u0440"),
                new TranslitException("Xorazm",     "\u0425\u043e\u0440\u0430\u0437\u043c",       "Geographic", "\u0412\u0438\u043b\u043e\u044f\u0442"),
                new TranslitException("Farg'ona",   "\u0424\u0430\u0440\u0493\u043e\u043d\u0430",      "Geographic", "\u0412\u0438\u043b\u043e\u044f\u0442"),

                // \u0418\u043c\u043b\u043e \u0432\u0430 \u043e\u0440\u0444\u043e\u0433\u0440\u0430\u0444\u0438\u044f \u0438\u0441\u0442\u0438\u0441\u043d\u043e\u043b\u0430\u0440\u0438 (\u0430\u043c\u0430\u043b\u0438\u0439)
                new TranslitException("sentabr",    "\u0441\u0435\u043d\u0442\u044f\u0431\u0440",      "Other", "\u0418\u043a\u043a\u0438 \u0451\u0437\u0443\u0432\u0434\u0430\u0433\u0438 \u043e\u0440\u0444\u043e\u0433\u0440\u0430\u0444\u0438\u044f \u0438\u0441\u0442\u0438\u0441\u043d\u043e\u0441\u0438"),
                new TranslitException("oktabr",     "\u043e\u043a\u0442\u044f\u0431\u0440",       "Other", "\u0418\u043a\u043a\u0438 \u0451\u0437\u0443\u0432\u0434\u0430\u0433\u0438 \u043e\u0440\u0444\u043e\u0433\u0440\u0430\u0444\u0438\u044f \u0438\u0441\u0442\u0438\u0441\u043d\u043e\u0441\u0438"),
                new TranslitException("MCHJ",       "\u041c\u0427\u0416",          "Other", "\u0410\u0431\u0431\u0440\u0435\u0432\u0438\u0430\u0442\u0443\u0440\u0430 \u0434\u043e\u0438\u043c \u043a\u0430\u0442\u0442\u0430 \u04b3\u0430\u0440\u0444\u043b\u0430\u0440\u0434\u0430"),
                new TranslitException("YEVB",       "\u0415\u0412\u0411",          "Other", "\u0410\u0431\u0431\u0440\u0435\u0432\u0438\u0430\u0442\u0443\u0440\u0430 \u0434\u043e\u0438\u043c \u043a\u0430\u0442\u0442\u0430 \u04b3\u0430\u0440\u0444\u043b\u0430\u0440\u0434\u0430"),
                new TranslitException("NATO",       "\u041d\u0410\u0422\u041e",         "Other", "\u0425\u0430\u043b\u049b\u0430\u0440\u043e \u0430\u0431\u0431\u0440\u0435\u0432\u0438\u0430\u0442\u0443\u0440\u0430"),
                new TranslitException("UNESCO",     "\u042e\u041d\u0415\u0421\u041a\u041e",       "Other", "\u0425\u0430\u043b\u049b\u0430\u0440\u043e \u0442\u0430\u0448\u043a\u0438\u043b\u043e\u0442 \u043d\u043e\u043c\u0438"),
                new TranslitException("COVID",      "\u041a\u041e\u0412\u0418\u0414",        "Other", "\u0422\u0438\u0431\u0431\u0438\u0439 \u0430\u0442\u0430\u043c\u0430"),
                new TranslitException("COVID-19",   "\u041a\u041e\u0412\u0418\u0414-19",     "Other", "\u0422\u0438\u0431\u0431\u0438\u0439 \u0430\u0442\u0430\u043c\u0430"),
            };
        }

        // =====================================================================
        //  PRIVATE HELPERS
        // =====================================================================


        /// <summary>
        /// Replaces exception words in text before applying standard transliteration.
        /// </summary>
        private string ApplyExceptions(string text, bool toCyrillic)
        {
            if (_exceptions == null || _exceptions.Count == 0) return text;

            foreach (var exc in _exceptions)
            {
                if (!exc.Enabled) continue;
                if (string.IsNullOrEmpty(exc.Original) || string.IsNullOrEmpty(exc.Replacement))
                    continue;

                // Determine source/target based on direction
                string find, replace;
                if (toCyrillic)
                {
                    // Converting to Cyrillic: exception.Original is Latin, Replacement is Cyrillic
                    find = exc.Original;
                    replace = exc.Replacement;
                }
                else
                {
                    // Converting to Latin: exception.Replacement is Cyrillic (was target), Original is Latin (target)
                    find = exc.Replacement;
                    replace = exc.Original;
                }

                // Case-insensitive whole-word replacement
                string pattern = @"\b" + Regex.Escape(find) + @"\b";
                text = Regex.Replace(text, pattern, match =>
                {
                    return PreserveCase(match.Value, replace);
                }, RegexOptions.IgnoreCase);
            }

            return text;
        }

        /// <summary>
        /// Handles Cyrillic Е at the beginning of words → "Ye" in Latin.
        /// Mid-word Е stays as "e".
        /// </summary>
        private string HandleWordInitialYe(string text)
        {
            // Word-initial Е → mark for Ye conversion
            // We use a temporary placeholder to avoid double-processing
            var result = new StringBuilder(text.Length + 20);
            bool atWordStart = true;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                if (c == 'Е' && atWordStart)
                {
                    result.Append("Йе"); // Will be caught by Й→Y, е→e  
                    atWordStart = false;
                }
                else if (c == 'е' && atWordStart)
                {
                    result.Append("йе"); // Will be caught by й→y, е→e
                    atWordStart = false;
                }
                else
                {
                    result.Append(c);
                    atWordStart = !char.IsLetter(c);
                }
            }

            return result.ToString();
        }

        /// <summary>
        /// Preserves the case pattern of the original word in the replacement.
        /// ALL CAPS → ALL CAPS, Title Case → Title Case, else lowercase.
        /// </summary>
        private static string PreserveCase(string original, string replacement)
        {
            if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(replacement))
                return replacement;

            if (original.All(c => !char.IsLetter(c) || char.IsUpper(c)))
                return replacement.ToUpperInvariant();

            if (char.IsUpper(original[0]))
            {
                if (replacement.Length == 1) return replacement.ToUpperInvariant();
                return char.ToUpperInvariant(replacement[0]) + replacement.Substring(1).ToLowerInvariant();
            }

            return replacement.ToLowerInvariant();
        }

#if DEBUG
        private void RunTransliterationSelfCheck()
        {
            try
            {
                var latinToCyr = new List<(string Latin, string CyrillicExpected)>
                {
                    ("o\u02BBzbek", "ўзбек"),
                    ("g\u02BBisht", "ғишт"),
                    ("shahar", "шаҳар"),
                    ("mas'ul", "масъул"),
                    ("sur'at", "суръат"),
                    ("qat'iy", "қатъий"),
                    ("'test'", "'тест'")
                };

                var cyrToLatin = new List<(string Cyrillic, string LatinExpected)>
                {
                    ("ўзбек", "o\u02BBzbek"),
                    ("ғишт", "g\u02BBisht"),
                    ("масъул", "mas\u02BBul"),
                    ("суръат", "sur\u02BBat"),
                    ("октябрь", "oktabr")
                };

                int failCount = 0;

                for (int i = 0; i < latinToCyr.Count; i++)
                {
                    var t = latinToCyr[i];
                    string actual = ToCyrillic(t.Latin);
                    if (!string.Equals(actual, t.CyrillicExpected, StringComparison.Ordinal))
                    {
                        failCount++;
                        Logger.Warn($"Translit self-check L2C mismatch: \"{t.Latin}\" -> \"{actual}\" (expected \"{t.CyrillicExpected}\")");
                    }
                }

                for (int i = 0; i < cyrToLatin.Count; i++)
                {
                    var t = cyrToLatin[i];
                    string actual = ToLatin(t.Cyrillic);
                    if (!string.Equals(actual, t.LatinExpected, StringComparison.Ordinal))
                    {
                        failCount++;
                        Logger.Warn($"Translit self-check C2L mismatch: \"{t.Cyrillic}\" -> \"{actual}\" (expected \"{t.LatinExpected}\")");
                    }
                }

                if (failCount == 0)
                    Logger.Info("Transliteration self-check passed.");
                else
                    Logger.Warn($"Transliteration self-check finished with {failCount} mismatches.");
            }
            catch (Exception ex)
            {
                Logger.Error("Transliteration self-check failed", ex);
            }
        }
#endif

        // =====================================================================
        //  SIMPLE JSON PARSER (no external dependencies for .NET 4.7.2)
        // =====================================================================

        /// <summary>
        /// Minimal JSON parser/serializer for TranslitException lists.
        /// Avoids dependency on Newtonsoft.Json or System.Text.Json.
        /// </summary>
        internal static class SimpleJsonParser
        {
            public static List<TranslitException> ParseExceptions(string json)
            {
                var result = new List<TranslitException>();
                if (string.IsNullOrWhiteSpace(json)) return result;

                json = json.Trim();
                if (!json.StartsWith("[")) return result;

                // Match each { ... } object
                var objectPattern = new Regex(@"\{[^}]*\}", RegexOptions.Singleline);
                var matches = objectPattern.Matches(json);

                foreach (Match m in matches)
                {
                    var obj = m.Value;
                    string original = ExtractJsonValue(obj, "Original")
                                   ?? ExtractJsonValue(obj, "original");
                    string replacement = ExtractJsonValue(obj, "Replacement")
                                      ?? ExtractJsonValue(obj, "replacement");
                    string category = ExtractJsonValue(obj, "Category")
                                   ?? ExtractJsonValue(obj, "category");
                    string description = ExtractJsonValue(obj, "Description")
                                      ?? ExtractJsonValue(obj, "description");
                    string note = ExtractJsonValue(obj, "Note")
                               ?? ExtractJsonValue(obj, "note");
                    string enabledStr = ExtractJsonValue(obj, "Enabled")
                                     ?? ExtractJsonValue(obj, "enabled");

                    // Parse Enabled — also handle bare true/false (not quoted)
                    bool enabled = true;
                    if (enabledStr != null)
                    {
                        enabled = !enabledStr.Equals("false", StringComparison.OrdinalIgnoreCase);
                    }
                    else
                    {
                        // Try bare boolean: "Enabled": false
                        var boolPat = new Regex(@"""Enabled""\s*:\s*(true|false)", RegexOptions.IgnoreCase);
                        var bm = boolPat.Match(obj);
                        if (bm.Success)
                            enabled = bm.Groups[1].Value.Equals("true", StringComparison.OrdinalIgnoreCase);
                    }

                    if (!string.IsNullOrEmpty(original))
                    {
                        result.Add(new TranslitException(original, replacement ?? "",
                            category ?? "Other", description, enabled, note));
                    }
                }

                return result;
            }

            public static string SerializeExceptions(List<TranslitException> exceptions)
            {
                if (exceptions == null || exceptions.Count == 0) return "[]";

                var sb = new StringBuilder();
                sb.AppendLine("[");

                for (int i = 0; i < exceptions.Count; i++)
                {
                    var e = exceptions[i];
                    sb.Append("  { ");
                    sb.Append($"\"Original\": \"{EscapeJson(e.Original)}\", ");
                    sb.Append($"\"Replacement\": \"{EscapeJson(e.Replacement)}\", ");
                    sb.Append($"\"Category\": \"{EscapeJson(e.Category ?? "Other")}\", ");
                    sb.Append($"\"Description\": \"{EscapeJson(e.Description ?? "")}\", ");
                    sb.Append($"\"Enabled\": {(e.Enabled ? "true" : "false")}, ");
                    sb.Append($"\"Note\": \"{EscapeJson(e.Note ?? "")}\"");
                    sb.Append(" }");
                    if (i < exceptions.Count - 1) sb.Append(",");
                    sb.AppendLine();
                }

                sb.Append("]");
                return sb.ToString();
            }

            private static string ExtractJsonValue(string json, string key)
            {
                // Pattern: "key" : "value"  or  "key": "value"
                var pattern = new Regex(
                    $"\"{Regex.Escape(key)}\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"",
                    RegexOptions.IgnoreCase);
                var match = pattern.Match(json);
                if (match.Success)
                {
                    return UnescapeJson(match.Groups[1].Value);
                }
                return null;
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
}
