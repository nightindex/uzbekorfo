using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// Grammar checking engine for Uzbek text — applies morphological,
    /// agreement, punctuation, and style rules.
    /// </summary>
    public interface IGrammarEngine
    {
        /// <summary>
        /// Checks a Word document range for grammar errors.
        /// </summary>
        List<ErrorEntry> CheckRange(Word.Range range);

        /// <summary>
        /// Checks plain text for grammar errors (no document context).
        /// </summary>
        List<ErrorEntry> CheckText(string text);

        /// <summary>
        /// Whether the grammar engine is currently active.
        /// </summary>
        bool IsEnabled { get; set; }

        /// <summary>
        /// Loads rule definitions from grammar_rules.json.
        /// </summary>
        void LoadRules();

        /// <summary>
        /// Enables or disables a specific rule.
        /// </summary>
        void SetRuleEnabled(string ruleId, bool enabled);

        /// <summary>
        /// Returns all loaded grammar rules.
        /// </summary>
        List<GrammarRule> GetRules();
    }
}
