using System.Collections.Generic;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// Predictive typing engine — provides inline word/phrase suggestions based on n-gram model.
    /// </summary>
    public interface IPredictionEngine
    {
        /// <summary>
        /// Returns prediction suggestions based on the preceding context text.
        /// </summary>
        List<string> GetPredictions(string precedingContext, int maxResults = 3);

        /// <summary>
        /// Learns vocabulary and n-gram patterns from the given text.
        /// Returns the number of new n-grams learned.
        /// </summary>
        int LearnFromText(string text);

        /// <summary>
        /// Saves the current prediction model to disk.
        /// </summary>
        void SaveModel();

        /// <summary>
        /// Loads the prediction model from disk.
        /// </summary>
        void LoadModel();

        /// <summary>
        /// Resets the prediction model to empty state.
        /// </summary>
        void ResetModel();

        /// <summary>
        /// Whether the prediction engine is currently active.
        /// </summary>
        bool IsEnabled { get; set; }

        /// <summary>
        /// Total number of learned n-gram entries.
        /// </summary>
        int ModelSize { get; }
    }
}
