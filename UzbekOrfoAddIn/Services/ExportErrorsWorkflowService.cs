using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates unresolved errors export flow (xlsx/xls/csv/txt).
    /// </summary>
    public sealed class ExportErrorsWorkflowService
    {
        private readonly ErrorStore _errorStore;

        public ExportErrorsWorkflowService(ErrorStore errorStore)
        {
            _errorStore = errorStore;
        }

        public ExportErrorsWorkflowResult ExecuteInteractive()
        {
            if (_errorStore == null || _errorStore.Count == 0)
                return ExportErrorsWorkflowResult.NoErrors();

            using (var dialog = BuildSaveDialog())
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return ExportErrorsWorkflowResult.Cancelled();

                var unresolved = _errorStore.Errors.Where(er => !er.IsResolved).ToList();
                if (unresolved.Count == 0)
                    return ExportErrorsWorkflowResult.NoUnresolvedErrors();

                if (dialog.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    ExportErrorsToXlsx(dialog.FileName, unresolved);
                    return ExportErrorsWorkflowResult.Completed(unresolved.Count, ExportErrorsFileKind.Xlsx);
                }

                if (dialog.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    ExportErrorsToXls(dialog.FileName, unresolved);
                    return ExportErrorsWorkflowResult.Completed(unresolved.Count, ExportErrorsFileKind.Xls);
                }

                var sb = new StringBuilder();
                bool isCsv = dialog.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase);
                if (isCsv)
                    sb.AppendLine("№,Сўз,Тури,Вариант,Параграф,Контекст");

                int i = 1;
                foreach (var err in unresolved)
                {
                    if (isCsv)
                    {
                        var errType = err.IsGrammarError ? "Грамматика" : "Имло";
                        sb.AppendLine($"{i},\"{err.Word}\",\"{errType}\",\"{err.BestSuggestion ?? string.Empty}\",{err.ParagraphIndex},\"{(err.Context?.Replace("\"", "'") ?? string.Empty)}\"");
                    }
                    else
                    {
                        var errType = err.IsGrammarError ? "[Грамматика]" : "[Имло]";
                        sb.AppendLine($"{i}. {err.Word} {errType} → {err.BestSuggestion ?? "?"} (§{err.ParagraphIndex})");
                    }
                    i++;
                }

                File.WriteAllText(dialog.FileName, sb.ToString(), Encoding.UTF8);
                return ExportErrorsWorkflowResult.Completed(
                    unresolved.Count,
                    isCsv ? ExportErrorsFileKind.Csv : ExportErrorsFileKind.Text);
            }
        }

        private static SaveFileDialog BuildSaveDialog()
        {
            return new SaveFileDialog
            {
                Title = "Хатоларни экспорт қилиш",
                Filter = "Excel файл (*.xlsx)|*.xlsx|Excel 97-2003 файл (*.xls)|*.xls|CSV файл (*.csv)|*.csv|Матн файл (*.txt)|*.txt",
                FileName = $"errors_{DateTime.Now:yyyyMMdd_HHmmss}"
            };
        }

        private static void ExportErrorsToXlsx(string filePath, List<ErrorEntry> errors)
        {
            if (errors == null) errors = new List<ErrorEntry>();
            XlsxExportHelper.WriteXlsxPackage(filePath, "Errors", BuildErrorsSheetXml(errors));
        }

        private static void ExportErrorsToXls(string filePath, List<ErrorEntry> errors)
        {
            string tempXlsx = Path.Combine(
                Path.GetTempPath(),
                "uzbekorfo_errors_" + Guid.NewGuid().ToString("N") + ".xlsx");

            try
            {
                ExportErrorsToXlsx(tempXlsx, errors);
                XlsxExportHelper.ConvertXlsxToXls(tempXlsx, filePath);
            }
            finally
            {
                try
                {
                    if (File.Exists(tempXlsx))
                        File.Delete(tempXlsx);
                }
                catch { }
            }
        }

        private static string BuildErrorsSheetXml(List<ErrorEntry> errors)
        {
            var sb = new StringBuilder(4096);
            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
            sb.AppendLine(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            sb.AppendLine(@"  <sheetViews><sheetView workbookViewId=""0""/></sheetViews>");
            sb.AppendLine(@"  <sheetFormatPr defaultRowHeight=""15""/>");
            sb.AppendLine(@"  <cols>");
            sb.AppendLine(@"    <col min=""1"" max=""1"" width=""6"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""2"" max=""2"" width=""24"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""3"" max=""3"" width=""14"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""4"" max=""4"" width=""24"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""5"" max=""5"" width=""12"" customWidth=""1""/>");
            sb.AppendLine(@"    <col min=""6"" max=""6"" width=""60"" customWidth=""1""/>");
            sb.AppendLine(@"  </cols>");
            sb.AppendLine(@"  <sheetData>");

            XlsxExportHelper.AppendInlineRow(sb, 1, new[] { "№", "Сўз", "Тури", "Вариант", "Параграф", "Контекст" });

            int row = 2;
            for (int i = 0; i < errors.Count; i++)
            {
                var err = errors[i];
                XlsxExportHelper.AppendInlineRow(sb, row++, new[]
                {
                    (i + 1).ToString(),
                    err.Word ?? string.Empty,
                    err.IsGrammarError ? "Грамматика" : "Имло",
                    err.BestSuggestion ?? string.Empty,
                    err.ParagraphIndex.ToString(),
                    err.Context ?? string.Empty
                });
            }

            sb.AppendLine(@"  </sheetData>");
            sb.AppendLine(@"</worksheet>");
            return sb.ToString();
        }
    }

    public sealed class ExportErrorsWorkflowResult
    {
        public ExportErrorsWorkflowStatus Status { get; private set; }
        public int ExportedCount { get; private set; }
        public ExportErrorsFileKind FileKind { get; private set; }

        public static ExportErrorsWorkflowResult Cancelled() =>
            new ExportErrorsWorkflowResult
            {
                Status = ExportErrorsWorkflowStatus.Cancelled
            };

        public static ExportErrorsWorkflowResult NoErrors() =>
            new ExportErrorsWorkflowResult
            {
                Status = ExportErrorsWorkflowStatus.NoErrors
            };

        public static ExportErrorsWorkflowResult NoUnresolvedErrors() =>
            new ExportErrorsWorkflowResult
            {
                Status = ExportErrorsWorkflowStatus.NoUnresolvedErrors
            };

        public static ExportErrorsWorkflowResult Completed(int exportedCount, ExportErrorsFileKind fileKind) =>
            new ExportErrorsWorkflowResult
            {
                Status = ExportErrorsWorkflowStatus.Completed,
                ExportedCount = Math.Max(0, exportedCount),
                FileKind = fileKind
            };
    }

    public enum ExportErrorsWorkflowStatus
    {
        Cancelled = 0,
        NoErrors = 1,
        NoUnresolvedErrors = 2,
        Completed = 3
    }

    public enum ExportErrorsFileKind
    {
        Text = 0,
        Csv = 1,
        Xls = 2,
        Xlsx = 3
    }
}
