using System.Drawing;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates "show definition for selected word" workflow.
    /// </summary>
    public sealed class DefinitionsWorkflowService
    {
        private readonly IExplanationProvider _provider;
        private readonly UzbekMorphAnalyzer _morphAnalyzer;
        private readonly ITransliterator _transliterator;

        public DefinitionsWorkflowService(
            IExplanationProvider provider,
            UzbekMorphAnalyzer morphAnalyzer = null,
            ITransliterator transliterator = null)
        {
            _provider = provider;
            _morphAnalyzer = morphAnalyzer;
            _transliterator = transliterator;
        }

        public DefinitionsWorkflowResult Execute(Image titleIcon)
        {
            if (!DocumentHelper.IsDocumentOpen())
                return DefinitionsWorkflowResult.NoOp();

            if (_provider == null)
                return DefinitionsWorkflowResult.ProviderUnavailable();

            string selectedText = DocumentHelper.GetSelectedWord();
            if (string.IsNullOrWhiteSpace(selectedText))
                return DefinitionsWorkflowResult.WordNotSelected();

            string word = selectedText.Trim();
            var entry = _provider.GetExplanation(word);
            string wordFormDetails = null;

            if (entry == null && _morphAnalyzer != null)
            {
                MorphAnalysis analysis = _morphAnalyzer.Analyze(word);
                if (analysis.IsKnownRoot)
                {
                    string root = _morphAnalyzer.GetDictionaryRoot(analysis);
                    entry = GetRootExplanation(root);
                    if (analysis.IsValidInflectedForm)
                        wordFormDetails = FormatWordFormDetails(analysis, root);
                }
            }

            var form = new ExplanationForm(word, entry, wordFormDetails);
            try { form.TitleIcon = titleIcon; } catch { }
            form.Show();

            return DefinitionsWorkflowResult.Completed(word);
        }

        private ExplanationEntry GetRootExplanation(string root)
        {
            if (string.IsNullOrWhiteSpace(root)) return null;

            ExplanationEntry entry = _provider.GetExplanation(root);
            if (entry != null || _transliterator == null) return entry;

            string latin = _transliterator.ToLatin(root);
            if (!string.IsNullOrWhiteSpace(latin) &&
                !latin.Equals(root, System.StringComparison.OrdinalIgnoreCase))
            {
                entry = _provider.GetExplanation(latin);
                if (entry != null) return entry;
            }

            string cyrillic = _transliterator.ToCyrillic(root);
            if (!string.IsNullOrWhiteSpace(cyrillic) &&
                !cyrillic.Equals(root, System.StringComparison.OrdinalIgnoreCase))
            {
                entry = _provider.GetExplanation(cyrillic);
            }

            return entry;
        }

        private string FormatWordFormDetails(MorphAnalysis analysis, string root)
        {
            if (analysis == null || !analysis.IsValidInflectedForm)
                return null;

            string displayRoot = root;
            var displaySuffixes = new System.Collections.Generic.List<string>(analysis.Suffixes);

            if (_transliterator != null && analysis.Script == ScriptType.Latin)
            {
                displayRoot = _transliterator.ToLatin(root);
                for (int i = 0; i < displaySuffixes.Count; i++)
                    displaySuffixes[i] = _transliterator.ToLatin(displaySuffixes[i]);
            }

            return "Асос: " + displayRoot + System.Environment.NewLine +
                   "Қўшимчалар: -" + string.Join(" + -", displaySuffixes);
        }
    }

    public sealed class DefinitionsWorkflowResult
    {
        public DefinitionsWorkflowStatus Status { get; private set; }
        public string Word { get; private set; }

        public static DefinitionsWorkflowResult NoOp() =>
            new DefinitionsWorkflowResult
            {
                Status = DefinitionsWorkflowStatus.NoOp
            };

        public static DefinitionsWorkflowResult ProviderUnavailable() =>
            new DefinitionsWorkflowResult
            {
                Status = DefinitionsWorkflowStatus.ProviderUnavailable
            };

        public static DefinitionsWorkflowResult WordNotSelected() =>
            new DefinitionsWorkflowResult
            {
                Status = DefinitionsWorkflowStatus.WordNotSelected
            };

        public static DefinitionsWorkflowResult Completed(string word) =>
            new DefinitionsWorkflowResult
            {
                Status = DefinitionsWorkflowStatus.Completed,
                Word = word
            };
    }

    public enum DefinitionsWorkflowStatus
    {
        NoOp = 0,
        ProviderUnavailable = 1,
        WordNotSelected = 2,
        Completed = 3
    }
}
