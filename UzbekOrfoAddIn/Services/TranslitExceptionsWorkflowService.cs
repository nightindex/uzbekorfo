using System.Drawing;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Forms;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates transliteration exceptions dialog setup and execution.
    /// </summary>
    public sealed class TranslitExceptionsWorkflowService
    {
        private readonly ITransliterator _transliterator;

        public TranslitExceptionsWorkflowService(ITransliterator transliterator)
        {
            _transliterator = transliterator;
        }

        public TranslitExceptionsWorkflowResult Execute(Image titleIcon)
        {
            if (_transliterator == null)
                return TranslitExceptionsWorkflowResult.ServiceUnavailable();

            var form = new TranslitExceptionsForm(
                getExceptions: () => _transliterator.GetExceptions(),
                onAdd: exc => _transliterator.AddException(exc),
                onUpdate: exc => _transliterator.AddException(exc),
                onRemove: original => _transliterator.RemoveException(original),
                onSave: () => _transliterator.SaveExceptions());

            try { form.TitleIcon = titleIcon; } catch { }
            form.ShowDialog();

            return TranslitExceptionsWorkflowResult.Completed();
        }
    }

    public sealed class TranslitExceptionsWorkflowResult
    {
        public TranslitExceptionsWorkflowStatus Status { get; private set; }

        public static TranslitExceptionsWorkflowResult ServiceUnavailable() =>
            new TranslitExceptionsWorkflowResult
            {
                Status = TranslitExceptionsWorkflowStatus.ServiceUnavailable
            };

        public static TranslitExceptionsWorkflowResult Completed() =>
            new TranslitExceptionsWorkflowResult
            {
                Status = TranslitExceptionsWorkflowStatus.Completed
            };
    }

    public enum TranslitExceptionsWorkflowStatus
    {
        ServiceUnavailable = 0,
        Completed = 1
    }
}
