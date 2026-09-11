using System.Drawing;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates "show definition for selected word" workflow.
    /// </summary>
    public sealed class DefinitionsWorkflowService
    {
        private readonly IExplanationProvider _provider;

        public DefinitionsWorkflowService(IExplanationProvider provider)
        {
            _provider = provider;
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

            var form = new ExplanationForm(word, entry);
            try { form.TitleIcon = titleIcon; } catch { }
            form.Show();

            return DefinitionsWorkflowResult.Completed(word);
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
