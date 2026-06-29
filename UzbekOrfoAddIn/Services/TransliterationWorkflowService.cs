using System;
using System.Windows.Forms;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.UI;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates transliteration commands for ribbon actions.
    /// </summary>
    public sealed class TransliterationWorkflowService
    {
        private readonly ITransliterator _transliterator;
        private readonly UzbekApostropheService _apostropheService;

        public TransliterationWorkflowService(
            ITransliterator transliterator,
            UzbekApostropheService apostropheService)
        {
            _transliterator = transliterator;
            _apostropheService = apostropheService ?? new UzbekApostropheService();
        }

        public TransliterationWorkflowResult ExecuteLatinToCyrillic()
        {
            if (!DocumentHelper.IsDocumentOpen())
                return TransliterationWorkflowResult.NoOp();

            if (_transliterator == null)
                return TransliterationWorkflowResult.ServiceUnavailable();

            var range = DocumentHelper.GetTargetRange();
            if (range == null)
                return TransliterationWorkflowResult.NoOp();

            string sourceText = range.Text;
            if (string.IsNullOrWhiteSpace(sourceText))
                return TransliterationWorkflowResult.TextNotFound();

            var script = _transliterator.DetectScript(sourceText);
            if (script == ScriptType.Cyrillic)
            {
                if (!ModernMessageBox.Confirm("Матн аллақачон кириллда. Бари бир ўтказайми?",
                        "Лотиндан Кириллга"))
                {
                    return TransliterationWorkflowResult.Cancelled();
                }
            }
            else
            {
                if (!ModernMessageBox.Confirm("Бутун документни Лотиндан Кириллга алмаштирасизми?",
                        "Лотиндан Кириллга"))
                {
                    return TransliterationWorkflowResult.Cancelled();
                }
            }

            string converted = _transliterator.ToCyrillic(sourceText);
            DocumentHelper.BeginUndoRecord("Лотиндан Кириллга");
            try
            {
                DocumentHelper.ReplaceRangeText(range, converted);
            }
            finally
            {
                DocumentHelper.EndUndoRecord();
            }

            int wordCount = CountWords(converted);
            return TransliterationWorkflowResult.Completed("Кирилл", wordCount);
        }

        public TransliterationWorkflowResult ExecuteCyrillicToLatin()
        {
            if (!DocumentHelper.IsDocumentOpen())
                return TransliterationWorkflowResult.NoOp();

            if (_transliterator == null)
                return TransliterationWorkflowResult.ServiceUnavailable();

            var range = DocumentHelper.GetTargetRange();
            if (range == null)
                return TransliterationWorkflowResult.NoOp();

            string sourceText = range.Text;
            if (string.IsNullOrWhiteSpace(sourceText))
                return TransliterationWorkflowResult.TextNotFound();

            var script = _transliterator.DetectScript(sourceText);
            if (script == ScriptType.Latin)
            {
                if (!ModernMessageBox.Confirm("Матн аллақачон лотинда. Бари бир ўтказайми?",
                        "Кириллдан Лотинга"))
                {
                    return TransliterationWorkflowResult.Cancelled();
                }
            }
            else
            {
                if (!ModernMessageBox.Confirm("Бутун документни Кириллдан Лотинга алмаштирасизми?",
                        "Кириллдан Лотинга"))
                {
                    return TransliterationWorkflowResult.Cancelled();
                }
            }

            string converted = _transliterator.ToLatin(sourceText);
            converted = _apostropheService.NormalizeText(converted);
            DocumentHelper.BeginUndoRecord("Кириллдан Лотинга");
            try
            {
                DocumentHelper.ReplaceRangeText(range, converted);
            }
            finally
            {
                DocumentHelper.EndUndoRecord();
            }

            int wordCount = CountWords(converted);
            return TransliterationWorkflowResult.Completed("Лотин", wordCount);
        }

        public TransliterationWorkflowResult ExecuteSwitchScript()
        {
            if (!DocumentHelper.IsDocumentOpen())
                return TransliterationWorkflowResult.NoOp();

            if (_transliterator == null)
                return TransliterationWorkflowResult.ServiceUnavailable();

            var range = DocumentHelper.GetTargetRange();
            if (range == null)
                return TransliterationWorkflowResult.NoOp();

            string sourceText = range.Text;
            if (string.IsNullOrWhiteSpace(sourceText))
                return TransliterationWorkflowResult.TextNotFound();

            var script = _transliterator.DetectScript(sourceText);
            string converted;
            string direction;

            switch (script)
            {
                case ScriptType.Cyrillic:
                    converted = _apostropheService.NormalizeText(_transliterator.ToLatin(sourceText));
                    direction = "Кирилл → Лотин";
                    break;
                case ScriptType.Latin:
                    converted = _transliterator.ToCyrillic(sourceText);
                    direction = "Лотин → Кирилл";
                    break;
                case ScriptType.Mixed:
                    var result = ModernMessageBox.ConfirmOrCancel(
                        "Матнда аралаш ёзув аниқланди.\nКириллга ўтказайми?",
                        "Ёзувни алмаштириш",
                        "Кириллга", "Лотинга", "Бекор");

                    if (result == DialogResult.Yes)
                    {
                        converted = _transliterator.ToCyrillic(sourceText);
                        direction = "Аралаш → Кирилл";
                    }
                    else if (result == DialogResult.No)
                    {
                        converted = _apostropheService.NormalizeText(_transliterator.ToLatin(sourceText));
                        direction = "Аралаш → Лотин";
                    }
                    else
                    {
                        return TransliterationWorkflowResult.Cancelled();
                    }
                    break;
                default:
                    return TransliterationWorkflowResult.ScriptUndetermined();
            }

            DocumentHelper.BeginUndoRecord("Ёзувни алмаштириш");
            try
            {
                DocumentHelper.ReplaceRangeText(range, converted);
            }
            finally
            {
                DocumentHelper.EndUndoRecord();
            }

            return TransliterationWorkflowResult.Completed(direction, CountWords(converted));
        }

        private static int CountWords(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            return text.Split(new[] { ' ', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }
    }

    public sealed class TransliterationWorkflowResult
    {
        public TransliterationWorkflowStatus Status { get; private set; }
        public string Direction { get; private set; }
        public int WordCount { get; private set; }

        public static TransliterationWorkflowResult NoOp() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.NoOp
            };

        public static TransliterationWorkflowResult ServiceUnavailable() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.ServiceUnavailable
            };

        public static TransliterationWorkflowResult TextNotFound() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.TextNotFound
            };

        public static TransliterationWorkflowResult ScriptUndetermined() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.ScriptUndetermined
            };

        public static TransliterationWorkflowResult Cancelled() =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.Cancelled
            };

        public static TransliterationWorkflowResult Completed(string direction, int wordCount) =>
            new TransliterationWorkflowResult
            {
                Status = TransliterationWorkflowStatus.Completed,
                Direction = direction,
                WordCount = Math.Max(0, wordCount)
            };
    }

    public enum TransliterationWorkflowStatus
    {
        NoOp = 0,
        ServiceUnavailable = 1,
        TextNotFound = 2,
        ScriptUndetermined = 3,
        Cancelled = 4,
        Completed = 5
    }
}
