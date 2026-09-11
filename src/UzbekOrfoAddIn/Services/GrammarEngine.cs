using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Grammar checking engine for Uzbek text.
    /// Pipeline: Tokenize в†’ MorphAnalyze в†’ SentenceSegment в†’ RuleApply в†’ Errors.
    /// Operates script-transparently (normalizes to Cyrillic, reports in original script).
    /// </summary>
    public class GrammarEngine : IGrammarEngine
    {
        private readonly DictionaryService _dictionary;
        private readonly ITransliterator _transliterator;
        private readonly UzbekMorphAnalyzer _morphAnalyzer;

        private List<GrammarRule> _rules = new List<GrammarRule>();
        private Dictionary<string, GrammarRule> _ruleById =
            new Dictionary<string, GrammarRule>(StringComparer.OrdinalIgnoreCase);
        private HashSet<string> _properNouns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private List<ProperNounPhrase> _properNounPhrases = new List<ProperNounPhrase>();

        private string _rulesPath;
        private string _properNounsPath;

        public bool IsEnabled { get; set; } = true;

        // ---- Subject pronouns for verb-agreement checking ----
        private static readonly Dictionary<string, string> PronounToAgreement = new Dictionary<string, string>
        {
            { "РјРµРЅ", "РјР°РЅ" }, { "men", "man" },
            { "СЃРµРЅ", "СЃР°РЅ" }, { "sen", "san" },
            { "Р±РёР·", "РјРёР·" }, { "biz", "miz" },
            { "СЃРёР·", "СЃРёР·" }, { "siz", "siz" },
        };

        // Regex for punctuation-related rules
        private static readonly Regex DoubleSpacePattern = new Regex(@"  +", RegexOptions.Compiled);
        private static readonly Regex CommaNoSpaceAfter = new Regex(@",(?! |\d)", RegexOptions.Compiled);
        private static readonly Regex SpaceBeforeComma = new Regex(@" ,", RegexOptions.Compiled);
        private static readonly Regex PeriodNoSpaceOrCap = new Regex(
            @"\.(\p{Ll})", RegexOptions.Compiled);
        private static readonly Regex SentenceStartPattern = new Regex(
            @"(?:^|[.!?]\s+)(\p{Ll})", RegexOptions.Compiled);

        // Brackets for matching
        private static readonly Dictionary<char, char> BracketPairs = new Dictionary<char, char>
        {
            { '(', ')' }, { '[', ']' }, { '{', '}' },
            { '\u00AB', '\u00BB' }
        };

        // Symmetric quote chars (same opener/closer)
        private static readonly HashSet<char> SymmetricQuotes = new HashSet<char>
        {
            '"'
        };

        public GrammarEngine(DictionaryService dictionary, ITransliterator transliterator,
            UzbekMorphAnalyzer morphAnalyzer, string rulesPath, string properNounsPath)
        {
            _dictionary = dictionary ?? throw new ArgumentNullException(nameof(dictionary));
            _transliterator = transliterator ?? throw new ArgumentNullException(nameof(transliterator));
            _morphAnalyzer = morphAnalyzer ?? throw new ArgumentNullException(nameof(morphAnalyzer));
            _rulesPath = rulesPath;
            _properNounsPath = properNounsPath;
        }

        // =====================================================================
        //  LOADING
        // =====================================================================

        public void LoadRules()
        {
            LoadRulesFromFile(_rulesPath);
            LoadProperNouns(_properNounsPath);
        }

        private void LoadRulesFromFile(string path)
        {
            try
            {
                _rules.Clear();
                _ruleById.Clear();

                if (!File.Exists(path))
                {
                    Logger.Warn($"Grammar rules file not found: {path}");
                    return;
                }

                string json = File.ReadAllText(path, System.Text.Encoding.UTF8);
                var serializer = new JavaScriptSerializer();
                var data = serializer.Deserialize<Dictionary<string, object>>(json);

                if (data == null || !data.ContainsKey("rules")) return;

                var ruleList = data["rules"] as object[];
                if (ruleList == null) return;

                foreach (var item in ruleList)
                {
                    var dict = item as Dictionary<string, object>;
                    if (dict == null) continue;

                    _rules.Add(new GrammarRule
                    {
                        RuleId = GetStr(dict, "ruleId"),
                        Category = GetStr(dict, "category"),
                        Priority = GetInt(dict, "priority"),
                        IsEnabled = GetBool(dict, "isEnabled"),
                        Severity = GetStr(dict, "severity"),
                        DescriptionUz = GetStr(dict, "description_uz"),
                        DescriptionEn = GetStr(dict, "description_en")
                    });
                }

                _rules = _rules.OrderBy(r => r.Priority).ToList();
                _ruleById = _rules
                    .Where(r => !string.IsNullOrWhiteSpace(r?.RuleId))
                    .GroupBy(r => r.RuleId, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
                Logger.Info($"Р“СЂР°РјРјР°С‚РёРєР° Т›РѕРёРґР°Р»Р°СЂРё СЋРєР»Р°РЅРґРё: {_rules.Count} С‚Р° Т›РѕРёРґР°");
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to load grammar rules", ex);
            }
        }

        private void LoadProperNouns(string path)
        {
            try
            {
                _properNouns.Clear();
                _properNounPhrases.Clear();

                if (!File.Exists(path))
                {
                    Logger.Warn($"Proper nouns file not found: {path}");
                    return;
                }

                string json = File.ReadAllText(path, System.Text.Encoding.UTF8);
                var serializer = new JavaScriptSerializer();
                var data = serializer.Deserialize<Dictionary<string, object>>(json);

                if (data == null || !data.ContainsKey("nouns")) return;

                var nouns = data["nouns"] as object[];
                if (nouns == null) return;

                foreach (var noun in nouns)
                {
                    var value = noun?.ToString()?.Trim();
                    if (string.IsNullOrWhiteSpace(value)) continue;

                    var parts = Regex.Split(value, @"\s+")
                        .Select(TextHelper.NormalizeWord)
                        .Where(p => !string.IsNullOrWhiteSpace(p))
                        .ToArray();

                    if (parts.Length > 1)
                    {
                        _properNounPhrases.Add(new ProperNounPhrase
                        {
                            Display = value,
                            Tokens = parts
                        });
                    }
                    else
                    {
                        _properNouns.Add(value);
                    }
                }

                Logger.Info($"РђС‚РѕТ›Р»Рё РѕС‚Р»Р°СЂ СЋРєР»Р°РЅРґРё: {_properNouns.Count + _properNounPhrases.Count} С‚Р°");
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to load proper nouns", ex);
            }
        }

        // =====================================================================
        //  PUBLIC CHECK METHODS
        // =====================================================================

        public List<ErrorEntry> CheckRange(Word.Range range)
        {
            if (range == null) return new List<ErrorEntry>();

            try
            {
                string text = range.Text;
                if (string.IsNullOrWhiteSpace(text)) return new List<ErrorEntry>();

                int rangeStart = range.Start;
                var doc = range.Document;

                var errors = CheckText(text);
                int baseParagraphIndex = 1;
                try
                {
                    baseParagraphIndex = Math.Max(1, doc.Range(0, rangeStart).Paragraphs.Count);
                }
                catch { }

                // Enrich each error with Word.Range and ParagraphIndex
                foreach (var error in errors)
                {
                    try
                    {
                        int absStart = rangeStart + error.StartIndex;
                        int absEnd = rangeStart + error.EndIndex;

                        // Clamp to document range
                        int docEnd = doc.Content.End;
                        if (absEnd > docEnd) absEnd = docEnd;
                        if (absStart >= absEnd) continue;

                        error.Range = doc.Range(absStart, absEnd);

                        // Compute absolute paragraph index by counting paragraph marks
                        // inside the checked range and offsetting from range start.
                        int paraOffset = 0;
                        for (int ci = 0; ci < error.StartIndex && ci < text.Length; ci++)
                        {
                            if (text[ci] == '\r') paraOffset++;
                        }
                        error.ParagraphIndex = baseParagraphIndex + paraOffset;

                        // Build a real context snippet (surrounding text, В±25 chars)
                        if (error.Message != null && (error.Context == null || error.Context == error.Message))
                        {
                            error.Context = GetContext(text, error.StartIndex, 25);
                        }
                    }
                    catch
                    {
                        // Range may be invalid if document changed
                    }
                }

                return errors;
            }
            catch (Exception ex)
            {
                Logger.Error("GrammarEngine.CheckRange failed", ex);
                return new List<ErrorEntry>();
            }
        }

        public List<ErrorEntry> CheckText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<ErrorEntry>();

            var errors = new List<ErrorEntry>();

            try
            {
                // Detect script of input text
                ScriptType inputScript = _transliterator.DetectScript(text);

                // Tokenize
                var tokens = TextHelper.Tokenize(text);

                // Morphologically analyze each token
                var analyses = new List<TokenAnalysis>();
                foreach (var token in tokens)
                {
                    if (TextHelper.ShouldSkipWord(token.Normalized))
                    {
                        analyses.Add(new TokenAnalysis { Token = token, Morph = null });
                        continue;
                    }

                    var morph = _morphAnalyzer.Analyze(token.Original);
                    analyses.Add(new TokenAnalysis { Token = token, Morph = morph });
                }

                var sentences = SegmentSentences(analyses, text);
                ApplyRulesByPriority(analyses, sentences, tokens, text, errors, inputScript);
            }
            catch (Exception ex)
            {
                Logger.Error("GrammarEngine.CheckText failed", ex);
            }

            return errors;
        }

        private void ApplyRulesByPriority(
            List<TokenAnalysis> analyses,
            List<List<TokenAnalysis>> sentences,
            List<WordToken> tokens,
            string text,
            List<ErrorEntry> errors,
            ScriptType inputScript)
        {
            foreach (var rule in _rules.Where(r => r != null && r.IsEnabled))
            {
                switch (rule.RuleId)
                {
                    case "DAT_ALLOMORPH":
                        foreach (var ta in analyses)
                        {
                            if (ta.Morph == null || !ta.Morph.IsAnalyzed) continue;
                            CheckDativeAllomorph(ta, errors, inputScript);
                        }
                        break;

                    case "LOC_VOICING":
                        foreach (var ta in analyses)
                        {
                            if (ta.Morph == null || !ta.Morph.IsAnalyzed) continue;
                            CheckLocativeVoicing(ta, errors, inputScript);
                        }
                        break;

                    case "ABL_VOICING":
                        foreach (var ta in analyses)
                        {
                            if (ta.Morph == null || !ta.Morph.IsAnalyzed) continue;
                            CheckAblativeVoicing(ta, errors, inputScript);
                        }
                        break;

                    case "DOUBLE_PLURAL":
                        foreach (var ta in analyses)
                        {
                            if (ta.Morph == null || !ta.Morph.IsAnalyzed) continue;
                            CheckDoublePlural(ta, errors, inputScript);
                        }
                        break;

                    case "POSS_MUTATION":
                        foreach (var ta in analyses)
                        {
                            if (ta.Morph == null || !ta.Morph.IsAnalyzed) continue;
                            CheckPossessiveMutation(ta, errors, inputScript);
                        }
                        break;

                    case "VERB_AGREEMENT":
                        foreach (var sentence in sentences)
                            CheckVerbAgreement(sentence, errors, inputScript);
                        break;

                    case "COPULA_SPACING":
                        foreach (var sentence in sentences)
                            CheckCopulaSpacing(sentence, text, errors, inputScript);
                        break;

                    case "SENT_CAPITALIZATION":
                        CheckSentenceCapitalization(text, tokens, errors, inputScript);
                        break;

                    case "COMMA_SPACE":
                        CheckCommaSpacing(text, errors, inputScript);
                        break;

                    case "PERIOD_SPACE":
                        CheckPeriodSpacing(text, errors, inputScript);
                        break;

                    case "DOUBLE_SPACE":
                        CheckDoubleSpaces(text, errors, inputScript);
                        break;

                    case "REPEATED_WORD":
                        CheckRepeatedWords(tokens, errors, inputScript);
                        break;

                    case "PROPER_NOUN_CAPS":
                        CheckProperNounCaps(text, tokens, errors, inputScript);
                        break;

                    case "BRACKET_MATCH":
                        CheckBracketMatching(text, errors, inputScript);
                        break;
                }
            }
        }

        public void SetRuleEnabled(string ruleId, bool enabled)
        {
            var rule = GetRule(ruleId);
            if (rule != null) rule.IsEnabled = enabled;
        }

        public List<GrammarRule> GetRules() => _rules.ToList();

        // =====================================================================
        //  MORPHOLOGICAL RULES
        // =====================================================================

        /// <summary>
        /// DAT_ALLOMORPH: checks -РіР°/-РєР°/-Т›Р° allomorph correctness.
        /// </summary>
        private void CheckDativeAllomorph(TokenAnalysis ta, List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("DAT_ALLOMORPH")) return;
            if (ta.Morph.SuffixIds == null) return;

            for (int i = 0; i < ta.Morph.SuffixIds.Count; i++)
            {
                string sid = ta.Morph.SuffixIds[i];
                if (sid != "DAT_GA" && sid != "DAT_KA" && sid != "DAT_QA") continue;
                if (!ta.Morph.IsKnownRoot) continue;

                string root = ta.Morph.Root;
                string expected = _morphAnalyzer.GetExpectedDativeSuffix(root);
                string actual = ta.Morph.Suffixes[i];

                if (actual != expected)
                {
                    string corrected = root + expected;
                    // Reconstruct with remaining suffixes
                    for (int j = i + 1; j < ta.Morph.Suffixes.Count; j++)
                        corrected += ta.Morph.Suffixes[j];

                    string displayCorrection = ConvertToOriginalScript(corrected, inputScript);
                    string message = $"В«{ta.Token.Original}В» вЂ” Р¶СћРЅР°Р»РёС€ РєРµР»РёС€РёРіРё РЅРѕС‚СћТ“СЂРё: " +
                                     $"-{ConvertToOriginalScript(actual, inputScript)} СћСЂРЅРёРіР° " +
                                     $"-{ConvertToOriginalScript(expected, inputScript)} Р±СћР»РёС€Рё РєРµСЂР°Рє";

                    errors.Add(CreateError(ta.Token, message, displayCorrection,
                        ErrorSeverity.Grammar, "DAT_ALLOMORPH"));
                }
            }
        }

        /// <summary>
        /// LOC_VOICING: checks -РґР°/-С‚Р° allomorph correctness.
        /// </summary>
        private void CheckLocativeVoicing(TokenAnalysis ta, List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("LOC_VOICING")) return;
            if (ta.Morph.SuffixIds == null) return;

            for (int i = 0; i < ta.Morph.SuffixIds.Count; i++)
            {
                string sid = ta.Morph.SuffixIds[i];
                if (sid != "LOC_DA" && sid != "LOC_TA") continue;
                if (!ta.Morph.IsKnownRoot) continue;

                string root = ta.Morph.Root;
                string expected = _morphAnalyzer.GetExpectedLocativeSuffix(root);
                string actual = ta.Morph.Suffixes[i];

                if (actual != expected)
                {
                    string corrected = root + expected;
                    for (int j = i + 1; j < ta.Morph.Suffixes.Count; j++)
                        corrected += ta.Morph.Suffixes[j];

                    string displayCorrection = ConvertToOriginalScript(corrected, inputScript);
                    string message = $"В«{ta.Token.Original}В» вЂ” СћСЂРёРЅ-РїР°Р№С‚ РєРµР»РёС€РёРіРё РЅРѕС‚СћТ“СЂРё: " +
                                     $"-{ConvertToOriginalScript(actual, inputScript)} СћСЂРЅРёРіР° " +
                                     $"-{ConvertToOriginalScript(expected, inputScript)}";

                    errors.Add(CreateError(ta.Token, message, displayCorrection,
                        ErrorSeverity.Grammar, "LOC_VOICING"));
                }
            }
        }

        /// <summary>
        /// ABL_VOICING: checks -РґР°РЅ/-С‚Р°РЅ allomorph correctness.
        /// </summary>
        private void CheckAblativeVoicing(TokenAnalysis ta, List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("ABL_VOICING")) return;
            if (ta.Morph.SuffixIds == null) return;

            for (int i = 0; i < ta.Morph.SuffixIds.Count; i++)
            {
                string sid = ta.Morph.SuffixIds[i];
                if (sid != "ABL_DAN" && sid != "ABL_TAN") continue;
                if (!ta.Morph.IsKnownRoot) continue;

                string root = ta.Morph.Root;
                string expected = _morphAnalyzer.GetExpectedAblativeSuffix(root);
                string actual = ta.Morph.Suffixes[i];

                if (actual != expected)
                {
                    string corrected = root + expected;
                    for (int j = i + 1; j < ta.Morph.Suffixes.Count; j++)
                        corrected += ta.Morph.Suffixes[j];

                    string displayCorrection = ConvertToOriginalScript(corrected, inputScript);
                    string message = $"В«{ta.Token.Original}В» вЂ” С‡РёТ›РёС€ РєРµР»РёС€РёРіРё РЅРѕС‚СћТ“СЂРё: " +
                                     $"-{ConvertToOriginalScript(actual, inputScript)} СћСЂРЅРёРіР° " +
                                     $"-{ConvertToOriginalScript(expected, inputScript)}";

                    errors.Add(CreateError(ta.Token, message, displayCorrection,
                        ErrorSeverity.Grammar, "ABL_VOICING"));
                }
            }
        }

        /// <summary>
        /// DOUBLE_PLURAL: detects patterns like -Р»Р°СЂ...Р»Р°СЂ (double plural).
        /// </summary>
        private void CheckDoublePlural(TokenAnalysis ta, List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("DOUBLE_PLURAL")) return;
            if (ta.Morph.SuffixIds == null) return;

            int pluralCount = ta.Morph.SuffixIds.Count(id => id == "PL" || id == "AGR_LAR");
            if (pluralCount >= 2)
            {
                string message = $"В«{ta.Token.Original}В» вЂ” РєСћРїР»РёРє Т›СћС€РёРјС‡Р°СЃРё С‚Р°РєСЂРѕСЂР»Р°РЅРіР°РЅ";
                errors.Add(CreateError(ta.Token, message, null,
                    ErrorSeverity.Grammar, "DOUBLE_PLURAL"));
            }
        }

        /// <summary>
        /// POSS_MUTATION: checks that Рєв†’Рі / Т›в†’Т“ mutation occurs before possessive suffixes.
        /// </summary>
        private void CheckPossessiveMutation(TokenAnalysis ta, List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("POSS_MUTATION")) return;
            if (ta.Morph.SuffixIds == null || ta.Morph.SuffixIds.Count == 0) return;

            // Find the first possessive suffix
            int possIdx = ta.Morph.SuffixIds.FindIndex(id => id.StartsWith("POSS_"));
            if (possIdx < 0) return;
            if (!ta.Morph.IsKnownRoot) return;

            if (_morphAnalyzer.HasMissingPossessiveMutation(ta.Morph.Root, ta.Morph.SuffixIds[possIdx]))
            {
                char last = ta.Morph.Root[ta.Morph.Root.Length - 1];
                char expected = last == '\u043A' ? '\u0433' : last == '\u049B' ? '\u0493' : last;

                string correctedRoot = ta.Morph.Root.Substring(0, ta.Morph.Root.Length - 1) + expected;
                string corrected = correctedRoot;
                foreach (var s in ta.Morph.Suffixes) corrected += s;

                string displayCorrection = ConvertToOriginalScript(corrected, inputScript);
                string message = $"В«{ta.Token.Original}В» вЂ” СЌРіР°Р»РёРє Т›СћС€РёРјС‡Р°СЃРё РѕР»РґРёРґР° " +
                                 $"СѓРЅРґРѕС€ СћР·РіР°СЂРёС€Рё: {last}в†’{expected}";

                errors.Add(CreateError(ta.Token, message, displayCorrection,
                    ErrorSeverity.Grammar, "POSS_MUTATION"));
            }
        }

        // =====================================================================
        //  AGREEMENT RULES (sentence-level)
        // =====================================================================

        /// <summary>
        /// VERB_AGREEMENT: checks that explicit pronoun subjects match verb person suffix.
        /// </summary>
        private void CheckVerbAgreement(List<TokenAnalysis> sentence, List<ErrorEntry> errors,
            ScriptType inputScript)
        {
            if (!IsRuleEnabled("VERB_AGREEMENT")) return;
            if (sentence.Count < 2) return;

            // Find explicit subject pronoun
            string pronoun = null;
            string expectedSuffix = null;

            foreach (var ta in sentence)
            {
                if (ta.Token == null) continue;
                string lower = ta.Token.Normalized;
                if (PronounToAgreement.ContainsKey(lower))
                {
                    pronoun = lower;
                    expectedSuffix = PronounToAgreement[lower];
                    break;
                }
            }

            if (pronoun == null || expectedSuffix == null) return;

            // Check the last token with verb morphology (SOV order)
            for (int i = sentence.Count - 1; i >= 0; i--)
            {
                var ta = sentence[i];
                if (ta.Morph == null || ta.Morph.PartOfSpeech != "verb") continue;

                // Check if the verb ends with the expected person suffix
                string cyrillicNorm = ta.Morph.NormalizedWord;
                if (ta.Morph.Script == ScriptType.Latin)
                    cyrillicNorm = _transliterator.ToCyrillic(ta.Morph.NormalizedWord);

                // Convert expected suffix to Cyrillic if needed
                string cyrExpected = expectedSuffix;

                if (!cyrillicNorm.EndsWith(cyrExpected, StringComparison.Ordinal))
                {
                    // Check if ANY agreement suffix is present (mismatch, not just missing)
                    bool hasAnyAgreement = ta.Morph.SuffixIds.Any(id => id.StartsWith("AGR_"));
                    if (hasAnyAgreement)
                    {
                        string message = $"В«{ta.Token.Original}В» вЂ” В«{pronoun}В» Р±РёР»Р°РЅ " +
                                         $"РјРѕСЃР»РёРє С…Р°С‚Рѕ, -{ConvertToOriginalScript(expectedSuffix, inputScript)} " +
                                         $"Р±СћР»РёС€Рё РєРµСЂР°Рє";
                        errors.Add(CreateError(ta.Token, message, null,
                            ErrorSeverity.Grammar, "VERB_AGREEMENT"));
                    }
                }
                break;
            }
        }

        /// <summary>
        /// COPULA_SPACING: detects "СЌРјР°СЃ РјР°РЅ" в†’ "СЌРјР°СЃРјР°РЅ" patterns.
        /// </summary>
        private void CheckCopulaSpacing(List<TokenAnalysis> sentence, string fullText,
            List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("COPULA_SPACING")) return;
            if (sentence.Count < 2) return;

            for (int i = 0; i < sentence.Count - 1; i++)
            {
                var first = sentence[i].Token;
                var second = sentence[i + 1].Token;
                if (first == null || second == null) continue;

                string w1 = first.Normalized;
                string w2 = second.Normalized;

                // Check for "СЌРјР°СЃ + person suffix" as separate words
                if (w1 == "СЌРјР°СЃ" || w1 == "emas")
                {
                    string w2lower = w2;
                    if (w2lower == "РјР°РЅ" || w2lower == "СЃР°РЅ" || w2lower == "РјРёР·" ||
                        w2lower == "СЃРёР·" || w2lower == "man" || w2lower == "san" ||
                        w2lower == "miz" || w2lower == "siz")
                    {
                        int spanStart = first.StartIndex;
                        int spanEnd = second.EndIndex;
                        if (spanEnd <= spanStart || spanEnd > fullText.Length) continue;
                        if (second.StartIndex <= first.EndIndex) continue;

                        string gap = fullText.Substring(first.EndIndex, second.StartIndex - first.EndIndex);
                        if (!string.IsNullOrWhiteSpace(gap)) continue;

                        string combined = first.Original + second.Original;
                        string phrase = fullText.Substring(spanStart, spanEnd - spanStart);
                        string message = $"В«{first.Original} {second.Original}В» вЂ” " +
                                         $"Р±РёСЂРіР° С‘Р·РёР»РёС€Рё РєРµСЂР°Рє: В«{combined}В»";
                        errors.Add(CreateSpanError(phrase, spanStart, spanEnd, message, combined,
                            ErrorSeverity.Grammar, "COPULA_SPACING"));
                    }
                }
            }
        }

        // =====================================================================
        //  PUNCTUATION & STYLE RULES
        // =====================================================================

        private void CheckSentenceCapitalization(string text, List<WordToken> tokens,
            List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("SENT_CAPITALIZATION")) return;
            if (tokens.Count == 0) return;

            // First word of text
            var first = tokens[0];
            if (first.Original.Length > 0 && char.IsLower(first.Original[0]))
            {
                string corrected = char.ToUpper(first.Original[0]) + first.Original.Substring(1);
                errors.Add(CreateError(first, "Р“Р°Рї Р±РѕС€РёРґР° Р±РѕС€ ТіР°СЂС„ С‘Р·РёР»РјР°РіР°РЅ", corrected,
                    ErrorSeverity.Punctuation, "SENT_CAPITALIZATION"));
            }

            // After sentence-ending punctuation
            for (int i = 1; i < tokens.Count; i++)
            {
                var prev = tokens[i - 1];
                var curr = tokens[i];

                // Check if there's sentence-ending punctuation between tokens
                int gapStart = prev.EndIndex;
                int gapEnd = curr.StartIndex;
                if (gapEnd <= gapStart || gapEnd > text.Length) continue;

                string gap = text.Substring(gapStart, gapEnd - gapStart);
                if (gap.Contains(".") || gap.Contains("!") || gap.Contains("?"))
                {
                    if (curr.Original.Length > 0 && char.IsLower(curr.Original[0]))
                    {
                        string corrected = char.ToUpper(curr.Original[0]) + curr.Original.Substring(1);
                        errors.Add(CreateError(curr,
                            "Р“Р°Рї Р±РѕС€РёРґР° Р±РѕС€ ТіР°СЂС„ С‘Р·РёР»РјР°РіР°РЅ", corrected,
                            ErrorSeverity.Punctuation, "SENT_CAPITALIZATION"));
                    }
                }
            }
        }

        private void CheckCommaSpacing(string text, List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("COMMA_SPACE")) return;
            var severity = GetRuleSeverity("COMMA_SPACE");
            var category = GetRule("COMMA_SPACE")?.Category ?? MapSeverityToCategory(severity);

            foreach (Match m in SpaceBeforeComma.Matches(text))
            {
                errors.Add(new ErrorEntry
                {
                    Word = " ,",
                    StartIndex = m.Index,
                    EndIndex = m.Index + m.Length,
                    Context = GetContext(text, m.Index, 20),
                    Message = "Р’РµСЂРіСѓР»РґР°РЅ РѕР»РґРёРЅ Р±СћС€Р»РёТ› Р±СћР»РјР°СЃР»РёРіРё РєРµСЂР°Рє",
                    Severity = severity,
                    RuleId = "COMMA_SPACE",
                    Category = category,
                    Suggestions = new List<Suggestion>
                    {
                        new Suggestion(",", 0.95, 1)
                    }
                });
            }

            foreach (Match m in CommaNoSpaceAfter.Matches(text))
            {
                if (m.Index + 1 < text.Length)
                {
                    errors.Add(new ErrorEntry
                    {
                        Word = text.Substring(m.Index, Math.Min(3, text.Length - m.Index)),
                        StartIndex = m.Index,
                        EndIndex = m.Index + m.Length,
                        Context = GetContext(text, m.Index, 20),
                        Message = "Р’РµСЂРіСѓР»РґР°РЅ РєРµР№РёРЅ Р±СћС€Р»РёТ› Р±СћР»РёС€Рё РєРµСЂР°Рє",
                        Severity = severity,
                        RuleId = "COMMA_SPACE",
                        Category = category,
                        Suggestions = new List<Suggestion>
                        {
                            new Suggestion(", ", 0.95, 1)
                        }
                    });
                }
            }
        }

        private void CheckPeriodSpacing(string text, List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("PERIOD_SPACE")) return;
            var severity = GetRuleSeverity("PERIOD_SPACE");
            var category = GetRule("PERIOD_SPACE")?.Category ?? MapSeverityToCategory(severity);

            foreach (Match m in PeriodNoSpaceOrCap.Matches(text))
            {
                errors.Add(new ErrorEntry
                {
                    Word = m.Value,
                    StartIndex = m.Index,
                    EndIndex = m.Index + m.Length,
                    Context = GetContext(text, m.Index, 20),
                    Message = "РќСѓТ›С‚Р°РґР°РЅ РєРµР№РёРЅ Р±СћС€Р»РёТ› РІР° Р±РѕС€ ТіР°СЂС„ Р±СћР»РёС€Рё РєРµСЂР°Рє",
                    Severity = severity,
                    RuleId = "PERIOD_SPACE",
                    Category = category,
                    Suggestions = new List<Suggestion>
                    {
                        new Suggestion(". " + char.ToUpper(m.Groups[1].Value[0]), 0.90, 1)
                    }
                });
            }
        }

        private void CheckDoubleSpaces(string text, List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("DOUBLE_SPACE")) return;
            var severity = GetRuleSeverity("DOUBLE_SPACE");
            var category = GetRule("DOUBLE_SPACE")?.Category ?? MapSeverityToCategory(severity);

            foreach (Match m in DoubleSpacePattern.Matches(text))
            {
                errors.Add(new ErrorEntry
                {
                    Word = m.Value,
                    StartIndex = m.Index,
                    EndIndex = m.Index + m.Length,
                    Context = GetContext(text, m.Index, 20),
                    Message = "РћСЂС‚РёТ›С‡Р° Р±СћС€Р»РёТ›Р»Р°СЂ С‚РѕРїРёР»РґРё",
                    Severity = severity,
                    RuleId = "DOUBLE_SPACE",
                    Category = category,
                    Suggestions = new List<Suggestion>
                    {
                        new Suggestion(" ", 0.99, m.Length - 1)
                    }
                });
            }
        }

        private void CheckRepeatedWords(List<WordToken> tokens, List<ErrorEntry> errors,
            ScriptType inputScript)
        {
            if (!IsRuleEnabled("REPEATED_WORD")) return;

            for (int i = 1; i < tokens.Count; i++)
            {
                if (string.Equals(tokens[i].Normalized, tokens[i - 1].Normalized,
                    StringComparison.OrdinalIgnoreCase) &&
                    !TextHelper.ShouldSkipWord(tokens[i].Normalized))
                {
                    string message = $"В«{tokens[i].Original}В» СЃСћР·Рё С‚Р°РєСЂРѕСЂР»Р°РЅРіР°РЅ";
                    errors.Add(CreateError(tokens[i], message, null,
                        ErrorSeverity.Style, "REPEATED_WORD"));
                }
            }
        }

        private void CheckProperNounCaps(string text, List<WordToken> tokens, List<ErrorEntry> errors,
            ScriptType inputScript)
        {
            if (!IsRuleEnabled("PROPER_NOUN_CAPS")) return;
            if (tokens == null || tokens.Count == 0) return;

            var seenSpans = new HashSet<string>(StringComparer.Ordinal);

            foreach (var token in tokens)
            {
                if (token.Original.Length < 2) continue;
                if (char.IsUpper(token.Original[0])) continue;

                // Check if this lowercased word is a known proper noun
                string capitalized = char.ToUpper(token.Original[0]) + token.Original.Substring(1);
                if (_properNouns.Contains(capitalized))
                {
                    string message = $"В«{token.Original}В» вЂ” Р°С‚РѕТ›Р»Рё РѕС‚, Р±РѕС€ ТіР°СЂС„ Р±РёР»Р°РЅ С‘Р·РёР»РёС€Рё РєРµСЂР°Рє";
                    string key = token.StartIndex + ":" + token.EndIndex;
                    if (!seenSpans.Add(key)) continue;

                    errors.Add(CreateError(token, message, capitalized,
                        ErrorSeverity.Grammar, "PROPER_NOUN_CAPS"));
                }
            }

            foreach (var phrase in _properNounPhrases)
            {
                if (phrase?.Tokens == null || phrase.Tokens.Length < 2) continue;
                int phraseLen = phrase.Tokens.Length;
                if (phraseLen > tokens.Count) continue;

                for (int i = 0; i <= tokens.Count - phraseLen; i++)
                {
                    bool match = true;
                    for (int j = 0; j < phraseLen; j++)
                    {
                        if (!string.Equals(tokens[i + j].Normalized, phrase.Tokens[j], StringComparison.OrdinalIgnoreCase))
                        {
                            match = false;
                            break;
                        }
                    }
                    if (!match) continue;

                    bool hasOnlyWhitespaceGaps = true;
                    for (int j = 0; j < phraseLen - 1; j++)
                    {
                        int gapStart = tokens[i + j].EndIndex;
                        int gapEnd = tokens[i + j + 1].StartIndex;
                        if (gapEnd < gapStart || gapEnd > text.Length)
                        {
                            hasOnlyWhitespaceGaps = false;
                            break;
                        }

                        string gap = text.Substring(gapStart, gapEnd - gapStart);
                        if (!string.IsNullOrWhiteSpace(gap))
                        {
                            hasOnlyWhitespaceGaps = false;
                            break;
                        }
                    }
                    if (!hasOnlyWhitespaceGaps) continue;

                    bool hasLowerCaseIssue = false;
                    for (int j = 0; j < phraseLen; j++)
                    {
                        var tok = tokens[i + j];
                        if (tok.Original.Length > 0 && char.IsLower(tok.Original[0]))
                        {
                            hasLowerCaseIssue = true;
                            break;
                        }
                    }
                    if (!hasLowerCaseIssue) continue;

                    int start = tokens[i].StartIndex;
                    int end = tokens[i + phraseLen - 1].EndIndex;
                    if (end <= start) continue;

                    string key = start + ":" + end;
                    if (!seenSpans.Add(key)) continue;

                    string observed = string.Join(" ", tokens.Skip(i).Take(phraseLen).Select(t => t.Original));
                    string message = $"В«{observed}В» вЂ” Р°С‚РѕТ›Р»Рё РѕС‚, Р±РѕС€ ТіР°СЂС„ Р±РёР»Р°РЅ С‘Р·РёР»РёС€Рё РєРµСЂР°Рє";
                    errors.Add(CreateSpanError(observed, start, end, message, phrase.Display,
                        ErrorSeverity.Grammar, "PROPER_NOUN_CAPS"));
                }
            }
        }

        private void CheckBracketMatching(string text, List<ErrorEntry> errors, ScriptType inputScript)
        {
            if (!IsRuleEnabled("BRACKET_MATCH")) return;

            var stack = new Stack<KeyValuePair<char, int>>();
            var closers = BracketPairs.Values.ToHashSet();
            var openerForCloser = BracketPairs.ToDictionary(kv => kv.Value, kv => kv.Key);

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (SymmetricQuotes.Contains(c))
                {
                    if (stack.Count > 0 && stack.Peek().Key == c)
                    {
                        stack.Pop();
                    }
                    else
                    {
                        stack.Push(new KeyValuePair<char, int>(c, i));
                    }
                }
                else if (BracketPairs.ContainsKey(c))
                {
                    stack.Push(new KeyValuePair<char, int>(c, i));
                }
                else if (closers.Contains(c))
                {
                    if (stack.Count > 0 && stack.Peek().Key == openerForCloser[c])
                    {
                        stack.Pop();
                    }
                    else
                    {
                        errors.Add(CreateSpanError(
                            c.ToString(),
                            i,
                            i + 1,
                            $"Ёпилувчи қавс «{c}» — жуфтсиз",
                            null,
                            ErrorSeverity.Punctuation,
                            "BRACKET_MATCH"));
                    }
                }
            }

            while (stack.Count > 0)
            {
                var unmatched = stack.Pop();
                errors.Add(CreateSpanError(
                    unmatched.Key.ToString(),
                    unmatched.Value,
                    unmatched.Value + 1,
                    $"Очилувчи қавс «{unmatched.Key}» — жуфтсиз",
                    null,
                    ErrorSeverity.Punctuation,
                    "BRACKET_MATCH"));
            }
        }

        // =====================================================================
        //  SENTENCE SEGMENTATION
        // =====================================================================

        private List<List<TokenAnalysis>> SegmentSentences(List<TokenAnalysis> analyses, string text)
        {
            var sentences = new List<List<TokenAnalysis>>();
            var current = new List<TokenAnalysis>();

            foreach (var ta in analyses)
            {
                current.Add(ta);

                if (ta.Token != null && ta.Token.EndIndex < text.Length)
                {
                    // Check for sentence-ending punctuation after this token
                    int afterIdx = ta.Token.EndIndex;
                    if (afterIdx < text.Length)
                    {
                        char afterChar = text[afterIdx];
                        if (afterChar == '.' || afterChar == '!' || afterChar == '?')
                        {
                            if (current.Count > 0)
                            {
                                sentences.Add(current);
                                current = new List<TokenAnalysis>();
                            }
                        }
                    }
                }
            }

            if (current.Count > 0)
                sentences.Add(current);

            return sentences;
        }

        // =====================================================================
        //  HELPERS
        // =====================================================================

        private GrammarRule GetRule(string ruleId)
        {
            if (string.IsNullOrWhiteSpace(ruleId)) return null;
            if (_ruleById != null && _ruleById.TryGetValue(ruleId, out var rule))
                return rule;
            return _rules.FirstOrDefault(r => string.Equals(r?.RuleId, ruleId, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsRuleEnabled(string ruleId)
        {
            var rule = GetRule(ruleId);
            return rule != null && rule.IsEnabled;
        }

        private ErrorSeverity GetRuleSeverity(string ruleId)
        {
            var rule = GetRule(ruleId);
            if (rule == null || string.IsNullOrWhiteSpace(rule.Severity))
                return ErrorSeverity.Grammar;

            switch (rule.Severity.Trim())
            {
                case "Grammar": return ErrorSeverity.Grammar;
                case "Punctuation": return ErrorSeverity.Punctuation;
                case "Style": return ErrorSeverity.Style;
                default: return ErrorSeverity.Grammar;
            }
        }

        private ErrorEntry CreateError(WordToken token, string message, string suggestion,
            ErrorSeverity severity, string ruleId)
        {
            return CreateSpanError(token.Original, token.StartIndex, token.EndIndex,
                message, suggestion, severity, ruleId);
        }

        private ErrorEntry CreateSpanError(string word, int startIndex, int endIndex,
            string message, string suggestion, ErrorSeverity severity, string ruleId)
        {
            var rule = GetRule(ruleId);
            var effectiveSeverity = rule != null ? GetRuleSeverity(ruleId) : severity;

            var entry = new ErrorEntry
            {
                Word = word,
                StartIndex = startIndex,
                EndIndex = endIndex,
                Context = message, // Will be replaced by real context in CheckRange
                Message = message,
                Severity = effectiveSeverity,
                RuleId = ruleId,
                Category = rule?.Category ?? MapSeverityToCategory(effectiveSeverity)
            };

            if (!string.IsNullOrEmpty(suggestion))
            {
                entry.Suggestions = new List<Suggestion>
                {
                    new Suggestion(suggestion, 0.95, 1)
                };
            }

            return entry;
        }

        private static string MapSeverityToCategory(ErrorSeverity severity)
        {
            switch (severity)
            {
                case ErrorSeverity.Grammar: return "morphological";
                case ErrorSeverity.Punctuation: return "punctuation";
                case ErrorSeverity.Style: return "style";
                default: return "grammar";
            }
        }

        // =====================================================================
        //  HIGHLIGHTING (green wavy underline)
        // =====================================================================

        /// <summary>
        /// Highlights grammar errors with green wavy underlines.
        /// </summary>
        public void HighlightErrors(List<ErrorEntry> errors)
        {
            if (errors == null) return;
            foreach (var error in errors)
                if (error != null && !error.IsResolved)
                    DocumentHighlightService.Highlight(error.Range, Word.WdColor.wdColorGreen);
        }

        /// <summary>
        /// Clears this session's temporary marks without touching existing underlines.
        /// </summary>
        public void ClearHighlights(Word.Document document)
        {
            DocumentHighlightService.ClearDocument(document);
        }

        private string ConvertToOriginalScript(string cyrillic, ScriptType targetScript)
        {
            if (targetScript == ScriptType.Latin && _transliterator != null)
                return _transliterator.ToLatin(cyrillic);
            return cyrillic;
        }

        private static string GetContext(string text, int position, int radius)
        {
            int start = Math.Max(0, position - radius);
            int end = Math.Min(text.Length, position + radius);
            return text.Substring(start, end - start);
        }

        private static string GetStr(Dictionary<string, object> dict, string key)
        {
            if (dict.ContainsKey(key) && dict[key] != null)
                return dict[key].ToString();
            return null;
        }

        private static int GetInt(Dictionary<string, object> dict, string key)
        {
            if (dict.ContainsKey(key) && dict[key] != null)
            {
                if (dict[key] is int i) return i;
                int.TryParse(dict[key].ToString(), out int result);
                return result;
            }
            return 0;
        }

        private static bool GetBool(Dictionary<string, object> dict, string key)
        {
            if (dict.ContainsKey(key) && dict[key] != null)
            {
                if (dict[key] is bool b) return b;
                bool.TryParse(dict[key].ToString(), out bool result);
                return result;
            }
            return false;
        }

        // =====================================================================
        //  INTERNAL TYPES
        // =====================================================================

        private class TokenAnalysis
        {
            public WordToken Token { get; set; }
            public MorphAnalysis Morph { get; set; }
        }

        private class ProperNounPhrase
        {
            public string Display { get; set; }
            public string[] Tokens { get; set; }
        }
    }
}

