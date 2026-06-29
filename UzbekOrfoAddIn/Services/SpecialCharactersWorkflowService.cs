using System.Drawing;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.UI;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates special-character insertion and full-document normalization.
    /// </summary>
    public sealed class SpecialCharactersWorkflowService
    {
        private readonly UzbekApostropheService _apostropheService;

        public SpecialCharactersWorkflowService(UzbekApostropheService apostropheService)
        {
            _apostropheService = apostropheService ?? new UzbekApostropheService();
        }

        public SpecialCharactersWorkflowResult InsertCharacter(string character)
        {
            if (!DocumentHelper.IsDocumentOpen())
                return SpecialCharactersWorkflowResult.NoOp();

            var sel = DocumentHelper.Selection;
            if (sel == null)
                return SpecialCharactersWorkflowResult.NoOp();

            sel.TypeText(character);
            return SpecialCharactersWorkflowResult.Completed(0);
        }

        public SpecialCharactersWorkflowResult NormalizeWholeDocument(Image confirmIcon)
        {
            if (!DocumentHelper.IsDocumentOpen())
                return SpecialCharactersWorkflowResult.NoOp();

            var doc = DocumentHelper.ActiveDoc;
            if (doc == null || doc.Content == null)
                return SpecialCharactersWorkflowResult.DocumentUnavailable();

            if (!ModernMessageBox.Confirm(
                "Бутун ҳужжатда тутуқ белгиларини автоматик тузатайми?",
                "Тутуқ белгисини тузатиш",
                "Ҳа, тузатиш",
                "Бекор",
                confirmIcon))
            {
                return SpecialCharactersWorkflowResult.Cancelled();
            }

            int corrected = _apostropheService.NormalizeRange(doc.Content, "Тутуқ белгисини тузатиш");
            return SpecialCharactersWorkflowResult.Completed(corrected);
        }
    }

    public sealed class SpecialCharactersWorkflowResult
    {
        public SpecialCharactersWorkflowStatus Status { get; private set; }
        public int CorrectedCount { get; private set; }

        public static SpecialCharactersWorkflowResult NoOp() =>
            new SpecialCharactersWorkflowResult
            {
                Status = SpecialCharactersWorkflowStatus.NoOp
            };

        public static SpecialCharactersWorkflowResult Cancelled() =>
            new SpecialCharactersWorkflowResult
            {
                Status = SpecialCharactersWorkflowStatus.Cancelled
            };

        public static SpecialCharactersWorkflowResult DocumentUnavailable() =>
            new SpecialCharactersWorkflowResult
            {
                Status = SpecialCharactersWorkflowStatus.DocumentUnavailable
            };

        public static SpecialCharactersWorkflowResult Completed(int correctedCount) =>
            new SpecialCharactersWorkflowResult
            {
                Status = SpecialCharactersWorkflowStatus.Completed,
                CorrectedCount = correctedCount
            };
    }

    public enum SpecialCharactersWorkflowStatus
    {
        NoOp = 0,
        Cancelled = 1,
        DocumentUnavailable = 2,
        Completed = 3
    }
}
