using System.Collections.Generic;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// Ranks spelling candidates with lexical/context/frequency signals
    /// and produces safe auto-correct decisions.
    /// </summary>
    public interface ISuggestionRanker
    {
        /// <summary>
        /// Current ranker version string for diagnostics.
        /// </summary>
        string RankerVersion { get; }

        /// <summary>
        /// Produces ranked suggestions from generated candidates.
        /// Input and candidates are expected in canonical (dictionary) script.
        /// </summary>
        List<Suggestion> RankCandidates(
            string originalWord,
            string canonicalInputWord,
            ScriptType inputScript,
            IList<CandidateMatch> candidates,
            SuggestionContext context,
            int maxResults);

        /// <summary>
        /// Applies conservative guardrails and chooses whether auto-correct should replace.
        /// </summary>
        AutoCorrectDecision BuildAutoCorrectDecision(
            string originalWord,
            List<Suggestion> rankedSuggestions,
            SuggestionContext context,
            string autoCorrectMode,
            double minAutoCorrectScore);

        /// <summary>
        /// Learns from accepted replacements to improve future ranking.
        /// </summary>
        void RecordAcceptedCorrection(string sourceWord, string replacementWord, SuggestionContext context);
    }
}
