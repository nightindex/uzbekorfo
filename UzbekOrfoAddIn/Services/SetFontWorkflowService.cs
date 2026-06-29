using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates font update workflow on selection or full document range.
    /// </summary>
    public sealed class SetFontWorkflowService
    {
        public SetFontWorkflowResult Execute(string fontName, float fontSize, string undoLabel)
        {
            if (!DocumentHelper.IsDocumentOpen())
                return SetFontWorkflowResult.NoOp();

            var range = DocumentHelper.GetTargetRange();
            if (range == null)
                return SetFontWorkflowResult.NoOp();

            DocumentHelper.BeginUndoRecord(undoLabel);
            try
            {
                range.Font.Name = fontName;
                range.Font.Size = fontSize;
            }
            finally
            {
                DocumentHelper.EndUndoRecord();
            }

            return SetFontWorkflowResult.Completed(fontName, fontSize);
        }
    }

    public sealed class SetFontWorkflowResult
    {
        public SetFontWorkflowStatus Status { get; private set; }
        public string FontName { get; private set; }
        public float FontSize { get; private set; }

        public static SetFontWorkflowResult NoOp() =>
            new SetFontWorkflowResult
            {
                Status = SetFontWorkflowStatus.NoOp
            };

        public static SetFontWorkflowResult Completed(string fontName, float fontSize) =>
            new SetFontWorkflowResult
            {
                Status = SetFontWorkflowStatus.Completed,
                FontName = fontName,
                FontSize = fontSize
            };
    }

    public enum SetFontWorkflowStatus
    {
        NoOp = 0,
        Completed = 1
    }
}
