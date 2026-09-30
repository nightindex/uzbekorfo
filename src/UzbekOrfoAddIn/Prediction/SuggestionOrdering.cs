using System;
using System.Collections.Generic;
using System.Linq;

namespace UzbekOrfoAddIn.Prediction
{
    // Shared by the Word adapter and corpus replay. Word validates safety and
    // morphology before merging; the replay measures ordering without that host step.
    internal static class SuggestionOrdering
    {
        private sealed class Option
        {
            internal string Text;
            internal double Score;
            internal int OriginalOrder;
        }

        internal static string[] Merge(string prefix, IEnumerable<PredictionCandidate> phrases,
            IEnumerable<string> lexical, int count, PredictionRankingPolicy policy)
        {
            if (count <= 0) return new string[0];
            var observed = (phrases ?? Enumerable.Empty<PredictionCandidate>()).ToArray();
            var dictionary = (lexical ?? Enumerable.Empty<string>()).ToArray();
            if (policy == PredictionRankingPolicy.Current || string.IsNullOrEmpty(prefix))
                return observed.Select(p => p.FullCompletion).Concat(dictionary)
                    .Where(s => !string.IsNullOrEmpty(s)).Distinct(StringComparer.Ordinal).Take(count).ToArray();

            // A dictionary word can outrank weak context evidence. Strong one- and
            // two-word contexts still take precedence over an isolated prefix match.
            var options = new List<Option>(observed.Length + dictionary.Length);
            for (int i = 0; i < observed.Length; i++)
                options.Add(new Option { Text = observed[i].FullCompletion, Score = observed[i].Score,
                    OriginalOrder = i });
            for (int i = 0; i < dictionary.Length; i++)
                options.Add(new Option { Text = dictionary[i], Score = 35 - i * 6,
                    OriginalOrder = observed.Length + i });
            return options.OrderByDescending(o => o.Score).ThenBy(o => o.OriginalOrder)
                .Select(o => o.Text).Where(s => !string.IsNullOrEmpty(s))
                .Distinct(StringComparer.Ordinal).Take(count).ToArray();
        }
    }
}
