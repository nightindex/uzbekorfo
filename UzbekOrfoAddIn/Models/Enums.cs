namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Detected script type of text content.
    /// </summary>
    public enum ScriptType
    {
        /// <summary>Unknown or insufficient text to detect.</summary>
        Unknown,

        /// <summary>Predominantly Latin Uzbek script.</summary>
        Latin,

        /// <summary>Predominantly Cyrillic Uzbek script.</summary>
        Cyrillic,

        /// <summary>Mixed Latin and Cyrillic characters.</summary>
        Mixed
    }

    /// <summary>
    /// Error severity classification.
    /// </summary>
    public enum ErrorSeverity
    {
        /// <summary>Definite spelling mistake.</summary>
        Spelling,

        /// <summary>Possible style or grammar issue.</summary>
        Style,

        /// <summary>Possible typo based on context.</summary>
        Typo,

        /// <summary>Grammar rule violation (morphological, agreement, etc.).</summary>
        Grammar,

        /// <summary>Punctuation or formatting issue.</summary>
        Punctuation
    }

    /// <summary>
    /// Result of adding a word to the dictionary.
    /// </summary>
    public enum AddWordResult
    {
        /// <summary>Word was successfully added.</summary>
        Added,

        /// <summary>Word already exists in the main dictionary.</summary>
        AlreadyInMainDictionary,

        /// <summary>Word already exists in the user dictionary.</summary>
        AlreadyInUserDictionary,

        /// <summary>Input was empty or invalid.</summary>
        Invalid
    }

    /// <summary>
    /// Export format options for error reports.
    /// </summary>
    public enum ExportFormat
    {
        /// <summary>Word document (.docx).</summary>
        Word,

        /// <summary>Plain text file (.txt).</summary>
        Text,

        /// <summary>Comma-separated values (.csv).</summary>
        Csv,

        /// <summary>HTML report (.html).</summary>
        Html,

        /// <summary>JSON data (.json).</summary>
        Json
    }
}
