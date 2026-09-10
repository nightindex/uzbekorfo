using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Builds a formatted Word report document for unresolved errors.
    /// </summary>
    public sealed class ErrorListWorkflowService
    {
        private readonly ErrorStore _errorStore;
        private readonly Word.Application _wordApplication;

        public ErrorListWorkflowService(ErrorStore errorStore, Word.Application wordApplication)
        {
            _errorStore = errorStore;
            _wordApplication = wordApplication;
        }

        public ErrorListWorkflowResult Execute(string documentName)
        {
            if (_errorStore == null || _errorStore.Count == 0)
                return ErrorListWorkflowResult.Empty();

            var unresolvedErrors = _errorStore.Errors.Where(er => !er.IsResolved).ToList();
            if (unresolvedErrors.Count == 0)
                return ErrorListWorkflowResult.AllResolved();

            if (_wordApplication == null)
                throw new InvalidOperationException("Word application instance is not available.");

            _wordApplication.ScreenUpdating = false;

            Word.Document reportDoc;
            try
            {
                reportDoc = _wordApplication.Documents.Add();
            }
            catch
            {
                _wordApplication.ScreenUpdating = true;
                throw;
            }

            try
            {
                BuildReport(reportDoc, unresolvedErrors, documentName);
            }
            finally
            {
                _wordApplication.ScreenUpdating = true;
            }

            return ErrorListWorkflowResult.Success(unresolvedErrors.Count);
        }

        private static void BuildReport(Word.Document reportDoc, System.Collections.Generic.List<Models.ErrorEntry> unresolvedErrors, string documentName)
        {
            var titleRange = reportDoc.Content;
            titleRange.Text = "ХАТОЛАР ҲИСОБОТИ\n";
            titleRange.Font.Name = "Times New Roman";
            titleRange.Font.Size = 18;
            titleRange.Font.Bold = 1;
            titleRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            titleRange.InsertParagraphAfter();

            int spellingCount = unresolvedErrors.Count(er => !er.IsGrammarError);
            int grammarCount = unresolvedErrors.Count(er => er.IsGrammarError);

            var metaRange = reportDoc.Paragraphs.Last.Range;
            metaRange.Font.Size = 11;
            metaRange.Font.Bold = 0;
            metaRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            metaRange.Text = $"Ҳужжат: {documentName}\n" +
                             $"Сана: {DateTime.Now:dd.MM.yyyy  HH:mm}\n" +
                             $"Жами хатолар: {unresolvedErrors.Count} та (имло: {spellingCount}, грамматика: {grammarCount})\n";
            metaRange.InsertParagraphAfter();

            var tableRange = reportDoc.Paragraphs.Last.Range;
            int rows = unresolvedErrors.Count + 1;
            int cols = 6;

            var table = reportDoc.Tables.Add(tableRange, rows, cols);
            table.Borders.Enable = 1;
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 11;

            table.Columns[1].Width = 28f;
            table.Columns[2].Width = 50f;
            table.Columns[3].Width = 85f;
            table.Columns[4].Width = 85f;
            table.Columns[5].Width = 45f;
            table.Columns[6].Width = 200f;

            var headerRow = table.Rows[1];
            headerRow.Range.Font.Bold = 1;
            headerRow.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray15;
            table.Cell(1, 1).Range.Text = "№";
            table.Cell(1, 2).Range.Text = "Тур";
            table.Cell(1, 3).Range.Text = "Сўз";
            table.Cell(1, 4).Range.Text = "Таклиф";
            table.Cell(1, 5).Range.Text = "§";
            table.Cell(1, 6).Range.Text = "Контекст / Қоида";

            for (int i = 0; i < unresolvedErrors.Count; i++)
            {
                var err = unresolvedErrors[i];
                int row = i + 2;
                bool isGrammar = err.IsGrammarError;

                table.Cell(row, 1).Range.Text = (i + 1).ToString();
                table.Cell(row, 2).Range.Text = isGrammar ? "Грамматика" : "Имло";
                table.Cell(row, 2).Range.Font.Size = 9;

                table.Cell(row, 3).Range.Text = err.Word ?? string.Empty;
                table.Cell(row, 3).Range.Font.Color = isGrammar
                    ? Word.WdColor.wdColorGreen
                    : Word.WdColor.wdColorRed;

                table.Cell(row, 4).Range.Text = err.BestSuggestion ?? "—";
                table.Cell(row, 5).Range.Text = $"§{err.ParagraphIndex}";
                table.Cell(row, 6).Range.Text = isGrammar
                    ? (err.Message ?? err.Context ?? string.Empty)
                    : (err.Context ?? string.Empty);
                table.Cell(row, 6).Range.Font.Size = 10;
            }

            var footerRange = reportDoc.Paragraphs.Last.Range;
            footerRange.InsertParagraphAfter();
            footerRange = reportDoc.Paragraphs.Last.Range;
            footerRange.Font.Size = 9;
            footerRange.Font.Italic = 1;
            footerRange.Font.Color = Word.WdColor.wdColorGray50;
            footerRange.Text = $"\nЎзбек Орфо — автоматик хатолар ҳисоботи | {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
        }
    }

    public sealed class ErrorListWorkflowResult
    {
        public ErrorListWorkflowStatus Status { get; private set; }
        public int GeneratedCount { get; private set; }

        public static ErrorListWorkflowResult Empty() =>
            new ErrorListWorkflowResult { Status = ErrorListWorkflowStatus.Empty, GeneratedCount = 0 };

        public static ErrorListWorkflowResult AllResolved() =>
            new ErrorListWorkflowResult { Status = ErrorListWorkflowStatus.AllResolved, GeneratedCount = 0 };

        public static ErrorListWorkflowResult Success(int generatedCount) =>
            new ErrorListWorkflowResult { Status = ErrorListWorkflowStatus.Success, GeneratedCount = generatedCount };
    }

    public enum ErrorListWorkflowStatus
    {
        Empty = 0,
        AllResolved = 1,
        Success = 2
    }
}
