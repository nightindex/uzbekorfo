using UzbekOrfoAddIn.Helpers;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates multi-step whitespace cleanup workflow.
    /// </summary>
    public sealed class CleanupSpacesWorkflowService
    {
        public CleanupSpacesWorkflowResult Execute()
        {
            if (!DocumentHelper.IsDocumentOpen())
                return CleanupSpacesWorkflowResult.NoOp();

            var doc = DocumentHelper.ActiveDoc;
            if (doc == null || doc.Content == null)
                return CleanupSpacesWorkflowResult.DocumentUnavailable();

            string logicalText = (doc.Content.Text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Trim();

            if (string.IsNullOrWhiteSpace(logicalText))
                return CleanupSpacesWorkflowResult.DocumentEmpty();

            DocumentHelper.BeginUndoRecord("Бўшлиқларни тозалаш");
            try
            {
                var find = doc.Content.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindContinue;
                find.MatchWildcards = true;

                find.Text = " {2,}";
                find.Replacement.Text = " ";
                find.Execute(Replace: Word.WdReplace.wdReplaceAll);

                find.Text = " {1,}^13";
                find.Replacement.Text = "^p";
                find.Execute(Replace: Word.WdReplace.wdReplaceAll);

                find.Text = "^13 {1,}";
                find.Replacement.Text = "^p";
                find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }
            finally
            {
                DocumentHelper.EndUndoRecord();
            }

            return CleanupSpacesWorkflowResult.Completed();
        }
    }

    public sealed class CleanupSpacesWorkflowResult
    {
        public CleanupSpacesWorkflowStatus Status { get; private set; }

        public static CleanupSpacesWorkflowResult NoOp() =>
            new CleanupSpacesWorkflowResult
            {
                Status = CleanupSpacesWorkflowStatus.NoOp
            };

        public static CleanupSpacesWorkflowResult DocumentUnavailable() =>
            new CleanupSpacesWorkflowResult
            {
                Status = CleanupSpacesWorkflowStatus.DocumentUnavailable
            };

        public static CleanupSpacesWorkflowResult DocumentEmpty() =>
            new CleanupSpacesWorkflowResult
            {
                Status = CleanupSpacesWorkflowStatus.DocumentEmpty
            };

        public static CleanupSpacesWorkflowResult Completed() =>
            new CleanupSpacesWorkflowResult
            {
                Status = CleanupSpacesWorkflowStatus.Completed
            };
    }

    public enum CleanupSpacesWorkflowStatus
    {
        NoOp = 0,
        DocumentUnavailable = 1,
        DocumentEmpty = 2,
        Completed = 3
    }
}
