using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using UzbekOrfoAddIn.Core;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates interactive dictionary import flow:
    /// file selection, dictionary import, and optional explanation merge.
    /// </summary>
    public sealed class ImportDictionaryWorkflowService
    {
        private readonly DictionaryService _dictionaryService;
        private readonly IExplanationProvider _explanationProvider;
        private static readonly string[] SupportedExtensions =
            { ".txt", ".dic", ".csv", ".doc", ".docx", ".xls", ".xlsx", ".json" };

        public ImportDictionaryWorkflowService(
            DictionaryService dictionaryService,
            IExplanationProvider explanationProvider)
        {
            _dictionaryService = dictionaryService;
            _explanationProvider = explanationProvider;
        }

        public ImportDictionaryWorkflowResult ExecuteInteractive()
        {
            using (var dialog = BuildImportDialog())
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return ImportDictionaryWorkflowResult.Cancelled();

                return Execute(dialog.FileName);
            }
        }

        public ImportDictionaryWorkflowResult ExecuteInteractiveFolder(SearchOption searchOption)
        {
            using (var dialog = new FolderBrowserDialog
            {
                Description = "Папкани танланг",
                ShowNewFolderButton = false
            })
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return ImportDictionaryWorkflowResult.Cancelled();

                return ExecuteFolder(dialog.SelectedPath, searchOption);
            }
        }

        public ImportDictionaryWorkflowResult Execute(string filePath)
        {
            if (_dictionaryService == null)
                return ImportDictionaryWorkflowResult.ServiceUnavailable();

            var import = _dictionaryService.ImportFromFileDetailed(filePath);
            int definitionsImported = 0;
            if (_explanationProvider != null &&
                import.DefinitionEntries != null &&
                import.DefinitionEntries.Count > 0)
            {
                definitionsImported = _explanationProvider.AddOrUpdateBatch(import.DefinitionEntries);
            }

            return ImportDictionaryWorkflowResult.Completed(
                import.AddedWordCount,
                definitionsImported,
                import.ProcessedFileCount,
                import.FailedFileCount,
                skippedInMain: import.SkippedInMainCount);
        }

        public ImportDictionaryWorkflowResult ExecuteFolder(string folderPath, SearchOption searchOption)
        {
            if (_dictionaryService == null)
                return ImportDictionaryWorkflowResult.ServiceUnavailable();

            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                return ImportDictionaryWorkflowResult.NoFiles();

            var files = Directory.EnumerateFiles(folderPath, "*.*", searchOption)
                .Where(IsSupportedExtension)
                .ToList();

            if (files.Count == 0)
                return ImportDictionaryWorkflowResult.NoFiles();

            var import = _dictionaryService.ImportFromFilesDetailed(files);
            int definitionsImported = 0;
            if (_explanationProvider != null &&
                import.DefinitionEntries != null &&
                import.DefinitionEntries.Count > 0)
            {
                definitionsImported = _explanationProvider.AddOrUpdateBatch(import.DefinitionEntries);
            }

            return ImportDictionaryWorkflowResult.Completed(
                import.AddedWordCount,
                definitionsImported,
                import.ProcessedFileCount,
                import.FailedFileCount,
                skippedInMain: import.SkippedInMainCount);
        }

        /// <summary>
        /// Collects the list of supported files from a folder. Returns null if folder is invalid.
        /// </summary>
        public List<string> CollectFolderFiles(string folderPath, SearchOption searchOption)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                return null;

            return Directory.EnumerateFiles(folderPath, "*.*", searchOption)
                .Where(IsSupportedExtension)
                .ToList();
        }

        /// <summary>
        /// Imports a pre-collected file list with progress reporting and cancellation support.
        /// Intended to be called from a BackgroundWorker.
        /// </summary>
        public ImportDictionaryWorkflowResult ExecuteFilesWithProgress(
            List<string> files,
            Action<int, int, string> progressCallback,
            Func<bool> isCancelled)
        {
            if (_dictionaryService == null)
                return ImportDictionaryWorkflowResult.ServiceUnavailable();

            if (files == null || files.Count == 0)
                return ImportDictionaryWorkflowResult.NoFiles();

            var import = _dictionaryService.ImportFromFilesDetailed(files, progressCallback, isCancelled);
            int definitionsImported = 0;
            if (_explanationProvider != null &&
                import.DefinitionEntries != null &&
                import.DefinitionEntries.Count > 0)
            {
                definitionsImported = _explanationProvider.AddOrUpdateBatch(import.DefinitionEntries);
            }

            return ImportDictionaryWorkflowResult.Completed(
                import.AddedWordCount,
                definitionsImported,
                import.ProcessedFileCount,
                import.FailedFileCount,
                skippedInMain: import.SkippedInMainCount);
        }

        internal static OpenFileDialog BuildImportDialog()
        {
            return new OpenFileDialog
            {
                Title = "Луғат файлини танланг",
                Filter =
                    "Қўллаб-қувватланадиган файллар (*.txt;*.dic;*.csv;*.doc;*.docx;*.xls;*.xlsx;*.json)|*.txt;*.dic;*.csv;*.doc;*.docx;*.xls;*.xlsx;*.json|" +
                    "Матн файллар (*.txt;*.dic;*.csv)|*.txt;*.dic;*.csv|" +
                    "Word ҳужжатлар (*.doc;*.docx)|*.doc;*.docx|" +
                    "Excel файллар (*.xls;*.xlsx)|*.xls;*.xlsx|" +
                    "JSON файллар (*.json)|*.json|" +
                    "Барча файллар (*.*)|*.*",
                FilterIndex = 1
            };
        }

        private static bool IsSupportedExtension(string path)
        {
            var ext = Path.GetExtension(path) ?? "";
            return SupportedExtensions.Contains(ext.ToLowerInvariant());
        }
    }

    public sealed class ImportDictionaryWorkflowResult
    {
        public ImportDictionaryWorkflowStatus Status { get; private set; }
        public int AddedWordCount { get; private set; }
        public int AddedDefinitionCount { get; private set; }
        public int ProcessedFileCount { get; private set; }
        public int FailedFileCount { get; private set; }
        public int SkippedInMainCount { get; private set; }
        public int TotalDictionaryWordCount { get; private set; }

        public static ImportDictionaryWorkflowResult Cancelled() =>
            new ImportDictionaryWorkflowResult
            {
                Status = ImportDictionaryWorkflowStatus.Cancelled
            };

        public static ImportDictionaryWorkflowResult ServiceUnavailable() =>
            new ImportDictionaryWorkflowResult
            {
                Status = ImportDictionaryWorkflowStatus.ServiceUnavailable
            };

        public static ImportDictionaryWorkflowResult NoFiles() =>
            new ImportDictionaryWorkflowResult
            {
                Status = ImportDictionaryWorkflowStatus.NoFiles
            };

        public static ImportDictionaryWorkflowResult Completed(
            int addedWords, int addedDefinitions, int processedFiles, int failedFiles,
            int totalDictionaryWords = 0, int skippedInMain = 0) =>
            new ImportDictionaryWorkflowResult
            {
                Status = ImportDictionaryWorkflowStatus.Completed,
                AddedWordCount = Math.Max(0, addedWords),
                AddedDefinitionCount = Math.Max(0, addedDefinitions),
                ProcessedFileCount = Math.Max(0, processedFiles),
                FailedFileCount = Math.Max(0, failedFiles),
                TotalDictionaryWordCount = Math.Max(0, totalDictionaryWords),
                SkippedInMainCount = Math.Max(0, skippedInMain)
            };
    }

    public enum ImportDictionaryWorkflowStatus
    {
        Cancelled = 0,
        ServiceUnavailable = 1,
        NoFiles = 2,
        Completed = 3
    }
}
