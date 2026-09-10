using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Services;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn
{
    public partial class UzbekOrfoRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        // Timer for debounced script indicator updates
        private Timer _scriptDetectTimer;

        // Cached script indicator icons to avoid re-creating bitmaps every 300ms
        private static readonly Font _iconFontSmall = new Font("Segoe UI", 9f, FontStyle.Bold);
        private static readonly Font _iconFontLarge = new Font("Segoe UI", 12f, FontStyle.Bold);
        private readonly Dictionary<string, Bitmap> _scriptIconCache = new Dictionary<string, Bitmap>();

        private void UzbekOrfoRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Logger.Info("UzbekOrfoRibbon loaded.");

            // Keep toggle UI aligned with persisted runtime state.
            var settings = ThisAddIn.Settings;
            if (settings != null)
            {
                toggleAutoCorrect.Checked = settings.AutoCorrectEnabled;
            }
            else
            {
                if (ThisAddIn.AutoCorrectService != null)
                    toggleAutoCorrect.Checked = ThisAddIn.AutoCorrectService.IsEnabled;
            }

            // OfficeImageId icons are not always exposed via RibbonButton.Image.
            // Prime the Definitions button image from Mso icon for form title usage.
            if (btnDefinitions != null && btnDefinitions.Image == null)
            {
                var thesaurusIcon = GetOfficeImageMso("Thesaurus", 32, 32);
                if (thesaurusIcon != null)
                    btnDefinitions.Image = thesaurusIcon;
            }

            // Script indicator: debounced timer and selection-change hook
            _scriptDetectTimer = new Timer { Interval = 300 };
            _scriptDetectTimer.Tick += (ts, te) => { _scriptDetectTimer.Stop(); UpdateScriptIndicator(); };

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app != null)
                {
                    app.WindowSelectionChange +=
                        new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(App_SelectionChangeForScriptIndicator);
                }
            }
            catch { }

            // Initial detection
            UpdateScriptIndicator();

            // Register keyboard hotkeys
            RegisterHotkeys();
        }

        // =================================================================
        //  SCRIPT INDICATOR
        // =================================================================

        private void App_SelectionChangeForScriptIndicator(Word.Selection sel)
        {
            try
            {
                _scriptDetectTimer?.Stop();
                _scriptDetectTimer?.Start();
            }
            catch { }
        }

        /// <summary>Detect dominant script in document and update ribbon label + colorful icon.</summary>
        private void UpdateScriptIndicator()
        {
            try
            {
                if (lblScriptIndicator == null) return;

                var app = Globals.ThisAddIn?.Application;
                var doc = app?.ActiveDocument;
                if (doc == null)
                {
                    SetScriptIndicatorState("\u2014", Color.FromArgb(120, 120, 120), "—");
                    return;
                }

                // Sample up to first 2000 chars for speed
                string sampleText = null;
                try
                {
                    var content = doc.Content;
                    if (content != null)
                    {
                        int len = content.End - content.Start;
                        if (len <= 0)
                        {
                            SetScriptIndicatorState("\u2014", Color.FromArgb(120, 120, 120), "—");
                            return;
                        }

                        int sampleEnd = Math.Min(content.Start + 2000, content.End);
                        var sampleRange = doc.Range(content.Start, sampleEnd);
                        sampleText = sampleRange.Text;
                    }
                }
                catch { }

                if (string.IsNullOrWhiteSpace(sampleText))
                {
                    SetScriptIndicatorState("\u2014", Color.FromArgb(120, 120, 120), "—");
                    return;
                }

                var translit = ThisAddIn.Transliterator;
                if (translit != null)
                {
                    var script = translit.DetectScript(sampleText);
                    switch (script)
                    {
                        case ScriptType.Cyrillic:
                            SetScriptIndicatorState("\u041A\u0438\u0440\u0438\u043B\u043B",
                                Color.FromArgb(59, 130, 246), "Ки");   // Blue
                            break;
                        case ScriptType.Latin:
                            SetScriptIndicatorState("\u041B\u043E\u0442\u0438\u043D",
                                Color.FromArgb(16, 185, 129), "Lt");   // Green
                            break;
                        case ScriptType.Mixed:
                            SetScriptIndicatorState("\u0410\u0440\u0430\u043B\u0430\u0448",
                                Color.FromArgb(245, 158, 11), "Mix");  // Amber
                            break;
                        default:
                            SetScriptIndicatorState("\u2014", Color.FromArgb(120, 120, 120), "—");
                            break;
                    }
                }
                else
                {
                    SetScriptIndicatorState("\u2014", Color.FromArgb(120, 120, 120), "—");
                }
            }
            catch
            {
                try { if (lblScriptIndicator != null) lblScriptIndicator.Label = "\u2014"; } catch { }
            }
        }

        /// <summary>Update the script indicator label and generate a colorful icon with text.</summary>
        private void SetScriptIndicatorState(string label, Color accentColor, string iconText)
        {
            lblScriptIndicator.Label = label;
            try
            {
                string cacheKey = accentColor.ToArgb() + "_" + iconText;
                if (!_scriptIconCache.TryGetValue(cacheKey, out var cachedIcon))
                {
                    cachedIcon = CreateScriptIndicatorIcon(accentColor, iconText);
                    _scriptIconCache[cacheKey] = cachedIcon;
                }
                lblScriptIndicator.Image = cachedIcon;
            }
            catch { }
        }

        /// <summary>
        /// Generate a 32x32 colorful ribbon icon: rounded-rect badge with accent color
        /// and short text (e.g. "Ки", "Lt", "Mix").
        /// </summary>
        private static Bitmap CreateScriptIndicatorIcon(Color accentColor, string text)
        {
            var bmp = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                // Rounded-rect background
                var rect = new Rectangle(1, 1, 30, 30);
                using (var path = new System.Drawing.Drawing2D.GraphicsPath())
                {
                    int r = 8;
                    path.AddArc(rect.X, rect.Y, r, r, 180, 90);
                    path.AddArc(rect.Right - r, rect.Y, r, r, 270, 90);
                    path.AddArc(rect.Right - r, rect.Bottom - r, r, r, 0, 90);
                    path.AddArc(rect.X, rect.Bottom - r, r, r, 90, 90);
                    path.CloseFigure();

                    using (var brush = new SolidBrush(accentColor))
                        g.FillPath(brush, path);
                }

                // White text centered
                var font = text.Length > 2 ? _iconFontSmall : _iconFontLarge;
                var textSize = g.MeasureString(text, font);
                float tx = (32 - textSize.Width) / 2f;
                float ty = (32 - textSize.Height) / 2f;
                using (var whiteBrush = new SolidBrush(Color.White))
                    g.DrawString(text, font, whiteBrush, tx, ty);
            }
            return bmp;
        }

        // =================================================================
        //  GROUP 1: ТЕКШИРИШ (Check)
        // =================================================================

        /// <summary>btnCheckSpelling — Имло текшируви (Spell Check)</summary>
        private void btnCheckSpelling_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                if (!DocumentHelper.IsDocumentOpen()) return;

                var engine = ThisAddIn.SpellingEngine;
                var store = ThisAddIn.ErrorStore;
                if (engine == null || store == null)
                {
                    SafeExecutor.ShowWarning("Имло текширув тизими юкланмаган.");
                    return;
                }

                // Get range to check: selection or entire document
                var range = DocumentHelper.GetTargetRange();
                if (range == null) return;

                var document = DocumentHelper.ActiveDoc;
                if (document == null) return;

                var workflow = new DocumentCheckWorkflowService(
                    engine,
                    ThisAddIn.GrammarEngine,
                    store);

                var result = workflow.Execute(
                    document,
                    range,
                    hasHighlights => ThisAddIn.SetSpellHighlightFlag(document, hasHighlights));

                if (result.TotalErrorCount == 0)
                {
                    ModernMessageBox.Success(
                        "Матн тўғри!\nИмло ва грамматик хатолар топилмади.",
                        "Матн текшируви",
                        this.btnCheckSpelling.Image);
                }
                else if (result.SpellingErrorCount > 0 && result.GrammarErrorCount > 0)
                {
                    // Both spelling and grammar errors
                    bool openErrors = ModernMessageBox.Confirm(
                        $"Имло хатолари: {result.SpellingErrorCount} та\n" +
                        $"Грамматик хатолар: {result.GrammarErrorCount} та\n\n" +
                        $"Жами: {result.TotalErrorCount} та хато топилди.",
                        "Матн текшируви",
                        "Хатоларни кўриш", "Тушундим",
                        this.btnCheckSpelling.Image);

                    if (openErrors) btnViewErrors_Click(sender, e);
                }
                else if (result.SpellingErrorCount > 0)
                {
                    // Only spelling errors
                    bool openErrors = ModernMessageBox.Confirm(
                        $"{result.SpellingErrorCount} та имло хатоси топилди.\n" +
                        "Грамматик хато йўқ.",
                        "Матн текшируви",
                        "Хатоларни кўриш", "Тушундим",
                        this.btnCheckSpelling.Image);

                    if (openErrors) btnViewErrors_Click(sender, e);
                }
                else
                {
                    // Only grammar errors
                    bool openErrors = ModernMessageBox.Confirm(
                        $"{result.GrammarErrorCount} та грамматик хато топилди.\n" +
                        "Имло хатоси йўқ.",
                        "Матн текшируви",
                        "Хатоларни кўриш", "Тушундим",
                        this.btnCheckSpelling.Image);

                    if (openErrors) btnViewErrors_Click(sender, e);
                }
            }, "Матн текшируви");
        }

        /// <summary>btnViewErrors — Хатоларни кўриш (View Errors)</summary>
        private void btnViewErrors_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                if (!DocumentHelper.IsDocumentOpen()) return;

                var workflow = new ViewErrorsWorkflowService(
                    ThisAddIn.ErrorStore,
                    word => ThisAddIn.AddWordToUserDictionary(word));

                workflow.Execute(this.btnViewErrors.Image);
            }, "Хатоларни кўриш");
        }

        /// <summary>btnErrorList — Хатолар рўйхати (Error List / formatted Word document report)</summary>
        private void btnErrorList_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new ErrorListWorkflowService(
                    ThisAddIn.ErrorStore,
                    Globals.ThisAddIn?.Application);

                var result = workflow.Execute(DocumentHelper.GetActiveDocumentName());
                switch (result.Status)
                {
                    case ErrorListWorkflowStatus.Empty:
                        SafeExecutor.ShowInfo("Хатолар рўйхати бўш. Аввал текширувни ишга туширинг.");
                        return;
                    case ErrorListWorkflowStatus.AllResolved:
                        SafeExecutor.ShowInfo("Барча хатолар тузатилган.");
                        return;
                    case ErrorListWorkflowStatus.Success:
                        ToastNotification.Success("Ҳисобот тайёр", $"{result.GeneratedCount} та хато рўйхати яратилди.");
                        return;
                }
            }, "Хатолар рўйхати");
        }

        // =================================================================
        //  GROUP 2: ТУЗАТИШ (Fix)
        // =================================================================

        /// <summary>btnSuggestions — Вариантлар (Suggestions for selected word)</summary>
        private void btnSuggestions_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                if (!DocumentHelper.IsDocumentOpen()) return;
                var workflow = new SuggestionsWorkflowService(
                    ThisAddIn.SpellingEngine,
                    ThisAddIn.GrammarEngine,
                    Globals.ThisAddIn?.Application,
                    word => ThisAddIn.AddWordToUserDictionary(word),
                    (fromWord, toWord) => ReplaceAllInDocument(fromWord, toWord));

                workflow.Execute(this.btnSuggestions.Image);
            }, "Вариантлар");
        }

        /// <summary>btnReplaceAll — Барчасини алмаштириш (Replace All)</summary>
        private void btnReplaceAll_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new ReplaceAllWorkflowService(ThisAddIn.ErrorStore);
                var analysis = workflow.Analyze();

                if (analysis.Status == ReplaceAllWorkflowStatus.EmptyStore)
                {
                    ModernMessageBox.Show(
                        "Тузатиладиган хато йўқ. Аввал текширувни ишга туширинг.",
                        "Барчасини алмаштириш",
                        ModernMessageBox.MessageType.Info,
                        "Тушундим",
                        this.btnReplaceAll.Image);
                    return;
                }

                if (analysis.Status == ReplaceAllWorkflowStatus.NoFixable)
                {
                    ModernMessageBox.Show(
                        "Автоматик тузатиш учун мос вариант йўқ.",
                        "Барчасини алмаштириш",
                        ModernMessageBox.MessageType.Info,
                        "Тушундим",
                        this.btnReplaceAll.Image);
                    return;
                }

                if (!ModernMessageBox.Confirm(
                        $"{analysis.FixableCount} та хатони автоматик тузатамизми?",
                        "Барчасини алмаштириш",
                        "Ҳа, тузатиш", "Бекор қилиш",
                        this.btnReplaceAll.Image))
                    return;

                var doc = DocumentHelper.ActiveDoc;
                if (doc == null)
                {
                    SafeExecutor.ShowInfo("Фаол ҳужжат топилмади.");
                    return;
                }

                int replaced = workflow.ApplyAll(doc, analysis.FixableErrors);

                ToastNotification.Success("Тузатилди", $"{replaced} та сўз алмаштирилди.");
            }, "Барчасини алмаштириш");
        }

        /// <summary>toggleAutoCorrect — Авто тузатиш (Auto-Correct toggle)</summary>
        private void toggleAutoCorrect_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                bool enabled = this.toggleAutoCorrect.Checked;

                // Persist setting
                var settings = ThisAddIn.Settings;
                if (settings != null)
                {
                    settings.AutoCorrectEnabled = enabled;
                    settings.Save();
                }

                // Enable/disable the real-time auto-correct service
                var autoCorrect = ThisAddIn.AutoCorrectService;
                if (autoCorrect != null)
                    autoCorrect.IsEnabled = enabled;

                if (enabled)
                {
                    ToastNotification.Success("Авто тузатиш", "Авто тузатиш ёқилди ✓");
                }
                else
                {
                    ToastNotification.Show("Авто тузатиш", "Авто тузатиш ўчирилди ✗",
                        ToastNotification.ToastType.Info);
                }
            }, "Авто тузатиш");
        }

        // =================================================================
        //  GROUP 3: ЛУҒАТ (Dictionary)
        // =================================================================

        /// <summary>btnAddToDict — Луғатга қўшиш (Add selected word to dictionary)</summary>
        private void btnAddToDict_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new AddToDictionaryWorkflowService(
                    (word, clearSelectionUnderline) =>
                        ThisAddIn.AddWordToUserDictionary(word, clearSelectionUnderline));

                var result = workflow.Execute();
                if (result.Status == AddToDictionaryWorkflowStatus.WordNotSelected)
                    SafeExecutor.ShowInfo("Аввал сўзни танланг.");
                else if (result.Status == AddToDictionaryWorkflowStatus.ServiceUnavailable)
                    SafeExecutor.ShowWarning("Луғат хизмати юкланмаган.");
            }, "Луғатга қўшиш");
        }

        /// <summary>btnEditDictionary — Луғатни таҳрирлаш (Edit Dictionary)</summary>
        private void btnEditDictionary_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new EditDictionaryWorkflowService(
                    ThisAddIn.DictionaryService,
                    ThisAddIn.ExplanationProvider);

                var result = workflow.Execute(btnEditDictionary.Image);
                if (result.Status == EditDictionaryWorkflowStatus.ServiceUnavailable)
                    SafeExecutor.ShowWarning("Луғат хизмати юкланмаган.");
            }, "Луғатни таҳрирлаш");
        }

        /// <summary>btnAddNewWords — Сўз қўшиш (Import words from file or folder)</summary>
        private void btnAddNewWords_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var dict = ThisAddIn.DictionaryService;
                if (dict == null) { SafeExecutor.ShowWarning("Луғат хизмати юкланмаган."); return; }

                var workflow = new ImportDictionaryWorkflowService(
                    dict,
                    ThisAddIn.ExplanationProvider);

                using (var dialog = new AddNewWordsForm())
                {
                    if (dialog.ShowDialog() != DialogResult.OK)
                        return;

                    ImportDictionaryWorkflowResult result;
                    if (dialog.SelectedMode == AddNewWordsMode.File)
                    {
                        result = ExecuteFileImportWithProgress(workflow);
                    }
                    else if (dialog.SelectedMode == AddNewWordsMode.Folder)
                    {
                        result = ExecuteFolderImportWithProgress(workflow);
                    }
                    else
                        return;

                    if (result == null || result.Status == ImportDictionaryWorkflowStatus.Cancelled)
                        return;
                    if (result.Status == ImportDictionaryWorkflowStatus.ServiceUnavailable)
                    {
                        SafeExecutor.ShowWarning("Луғат хизмати юкланмаган.");
                        return;
                    }
                    if (result.Status == ImportDictionaryWorkflowStatus.NoFiles)
                    {
                        ImportResultForm.ShowResult(0, 0, 0, dict.TotalWordCount);
                        return;
                    }

                    ImportResultForm.ShowResult(
                        result.AddedWordCount,
                        result.AddedDefinitionCount,
                        result.ProcessedFileCount,
                        dict.TotalWordCount,
                        result.SkippedInMainCount);
                }
            }, "Сўз қўшиш");
        }

        /// <summary>
        /// Shows a folder picker, then imports all supported files with a progress dialog.
        /// </summary>
        private ImportDictionaryWorkflowResult ExecuteFolderImportWithProgress(
            ImportDictionaryWorkflowService workflow)
        {
            // 1. Pick folder
            string folderPath;
            using (var folderDlg = new FolderBrowserDialog
            {
                Description = "Папкани танланг",
                ShowNewFolderButton = false
            })
            {
                if (folderDlg.ShowDialog() != DialogResult.OK)
                    return ImportDictionaryWorkflowResult.Cancelled();
                folderPath = folderDlg.SelectedPath;
            }

            // 2. Collect files
            var files = workflow.CollectFolderFiles(folderPath, SearchOption.AllDirectories);
            if (files == null || files.Count == 0)
                return ImportDictionaryWorkflowResult.NoFiles();

            // 3. Run with progress dialog
            ImportDictionaryWorkflowResult importResult = null;

            using (var progressForm = new ImportProgressForm())
            {
                var dr = progressForm.RunWithWork((progress) =>
                {
                    importResult = workflow.ExecuteFilesWithProgress(
                        files,
                        progressCallback: (current, total, fileName) =>
                        {
                            int pct = (int)((double)current / total * 100);
                            progress.ReportProgress(
                                pct,
                                $"Файл ишланмоқда: {current}/{total}",
                                fileName);
                        },
                        isCancelled: () => progress.CancellationPending);
                });

                if (dr == DialogResult.Cancel)
                    return ImportDictionaryWorkflowResult.Cancelled();

                if (dr == DialogResult.Abort || progressForm.WorkerError != null)
                {
                    Logger.Error("Папка импортида хато", progressForm.WorkerError);
                    SafeExecutor.ShowWarning("Импорт вақтида хато юз берди.");
                    return ImportDictionaryWorkflowResult.Cancelled();
                }
            }

            return importResult ?? ImportDictionaryWorkflowResult.NoFiles();
        }

        /// <summary>
        /// Shows a file picker, then imports a single file with a progress dialog.
        /// </summary>
        private ImportDictionaryWorkflowResult ExecuteFileImportWithProgress(
            ImportDictionaryWorkflowService workflow)
        {
            // 1. Pick file
            string filePath;
            using (var fileDlg = ImportDictionaryWorkflowService.BuildImportDialog())
            {
                if (fileDlg.ShowDialog() != DialogResult.OK)
                    return ImportDictionaryWorkflowResult.Cancelled();
                filePath = fileDlg.FileName;
            }

            if (string.IsNullOrWhiteSpace(filePath))
                return ImportDictionaryWorkflowResult.NoFiles();

            // 2. Run with progress dialog
            ImportDictionaryWorkflowResult importResult = null;

            using (var progressForm = new ImportProgressForm())
            {
                var dr = progressForm.RunWithWork((progress) =>
                {
                    var files = new List<string> { filePath };
                    progress.ReportProgress(5, "Импорт бошланмоқда...", Path.GetFileName(filePath));
                    importResult = workflow.ExecuteFilesWithProgress(
                        files,
                        progressCallback: (current, total, fileName) =>
                        {
                            int pct = total <= 1
                                ? 60
                                : (int)((double)current / total * 100);
                            progress.ReportProgress(
                                pct,
                                $"Файл ишланмоқда: {current}/{total}",
                                fileName);
                        },
                        isCancelled: () => progress.CancellationPending);
                    progress.ReportProgress(100, "Импорт тугади", Path.GetFileName(filePath));
                });

                if (dr == DialogResult.Cancel)
                    return ImportDictionaryWorkflowResult.Cancelled();

                if (dr == DialogResult.Abort || progressForm.WorkerError != null)
                {
                    Logger.Error("Файл импортида хато", progressForm.WorkerError);
                    SafeExecutor.ShowWarning("Импорт вақтида хато юз берди.");
                    return ImportDictionaryWorkflowResult.Cancelled();
                }
            }

            return importResult ?? ImportDictionaryWorkflowResult.NoFiles();
        }

        // =================================================================
        //  GROUP 4: ТРАНСЛИТЕРАЦИЯ (Transliteration)
        // =================================================================

        /// <summary>btnFromLatinToCyrillic — Лотиндан Кириллга (Latin → Cyrillic)</summary>
        private void btnFromLatinToCyrillic_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new TransliterationWorkflowService(
                    ThisAddIn.Transliterator,
                    new UzbekApostropheService());

                var result = workflow.ExecuteLatinToCyrillic();
                if (result.Status == TransliterationWorkflowStatus.ServiceUnavailable)
                {
                    SafeExecutor.ShowWarning("Транслитерация хизмати юкланмаган.");
                    return;
                }
                if (result.Status == TransliterationWorkflowStatus.TextNotFound)
                {
                    SafeExecutor.ShowInfo("Матн топилмади.");
                    return;
                }
                if (result.Status == TransliterationWorkflowStatus.Completed)
                    ToastNotification.Success("Транслитерация тугади", $"{result.WordCount} сўз кириллга ўтказилди.");
            }, "Лотиндан Кириллга");
        }

        /// <summary>btnFromCyrillicToLatin — Кириллдан Лотинга (Cyrillic → Latin)</summary>
        private void btnFromCyrillicToLatin_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new TransliterationWorkflowService(
                    ThisAddIn.Transliterator,
                    new UzbekApostropheService());

                var result = workflow.ExecuteCyrillicToLatin();
                if (result.Status == TransliterationWorkflowStatus.ServiceUnavailable)
                {
                    SafeExecutor.ShowWarning("Транслитерация хизмати юкланмаган.");
                    return;
                }
                if (result.Status == TransliterationWorkflowStatus.TextNotFound)
                {
                    SafeExecutor.ShowInfo("Матн топилмади.");
                    return;
                }
                if (result.Status == TransliterationWorkflowStatus.Completed)
                    ToastNotification.Success("Транслитерация тугади", $"{result.WordCount} сўз лотинга ўтказилди.");
            }, "Кириллдан Лотинга");
        }

        /// <summary>btnTransExceptions — Истиснолар (Manage transliteration exceptions)</summary>
        private void btnTransExceptions_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new TranslitExceptionsWorkflowService(ThisAddIn.Transliterator);
                var result = workflow.Execute(btnTransExceptions.Image);
                if (result.Status == TranslitExceptionsWorkflowStatus.ServiceUnavailable)
                    SafeExecutor.ShowWarning("Транслитерация хизмати юкланмаган.");
            }, "Истиснолар");
        }

        // =================================================================
        //  GROUP 5: ВОСИТАЛАР (Tools)
        // =================================================================

        /// <summary>btnCleanupSpaces — Бўшлиқларни тозалаш</summary>
        private void btnCleanupSpaces_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new CleanupSpacesWorkflowService();
                var result = workflow.Execute();
                if (result.Status == CleanupSpacesWorkflowStatus.DocumentUnavailable)
                {
                    SafeExecutor.ShowInfo("Фаол ҳужжат топилмади.");
                    return;
                }
                if (result.Status == CleanupSpacesWorkflowStatus.DocumentEmpty)
                {
                    ToastNotification.ShowInfo("Тозалаш", "Ҳужжат бўш.");
                    return;
                }
                if (result.Status == CleanupSpacesWorkflowStatus.Completed)
                    ToastNotification.Success("Тозаланди", "Ортиқча бўшлиқлар олиб ташланди.");
            }, "Бўшлиқларни тозалаш");
        }

        /// <summary>btnSwitchScript — Ёзувни алмаштириш (auto-detect and switch script)</summary>
        private void btnSwitchScript_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new TransliterationWorkflowService(
                    ThisAddIn.Transliterator,
                    new UzbekApostropheService());

                var result = workflow.ExecuteSwitchScript();
                if (result.Status == TransliterationWorkflowStatus.ServiceUnavailable)
                {
                    SafeExecutor.ShowWarning("Транслитерация хизмати юкланмаган.");
                    return;
                }
                if (result.Status == TransliterationWorkflowStatus.TextNotFound)
                {
                    SafeExecutor.ShowInfo("Матн топилмади.");
                    return;
                }
                if (result.Status == TransliterationWorkflowStatus.ScriptUndetermined)
                {
                    SafeExecutor.ShowInfo("Матнда ёзув тури аниқланмади.");
                    return;
                }
                if (result.Status == TransliterationWorkflowStatus.Completed)
                    ToastNotification.Success("Ўтказилди", result.Direction);
            }, "Ёзувни алмаштириш");
        }

        /// <summary>btnSetFontTNR — TNR ўрнатиш (set font to Times New Roman)</summary>
        private void btnSetFontTNR_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new SetFontWorkflowService();
                var result = workflow.Execute("Times New Roman", 14f, "TNR ўрнатиш");
                if (result.Status == SetFontWorkflowStatus.Completed)
                    ToastNotification.Success("Шрифт ўзгартирилди", "Times New Roman 14pt");
            }, "TNR ўрнатиш");
        }

        // =================================================================
        //  GROUP 6: ИЗОҲЛАР (Explanations)
        // =================================================================

        /// <summary>btnDefinitions — Изоҳ (Definition / Explanation for selected word)</summary>
        private void btnDefinitions_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new DefinitionsWorkflowService(ThisAddIn.ExplanationProvider);
                var result = workflow.Execute(ResolveDefinitionsTitleIcon());
                if (result.Status == DefinitionsWorkflowStatus.ProviderUnavailable)
                {
                    SafeExecutor.ShowWarning("Изоҳлар базаси юкланмаган ёки бўш.");
                    return;
                }
                if (result.Status == DefinitionsWorkflowStatus.WordNotSelected)
                {
                    SafeExecutor.ShowInfo("Аввал сўзни танланг.");
                }
            }, "Изоҳ");
        }

        /// <summary>btnExportErrors — Экспорт (Export errors to file)</summary>
        private void btnExportErrors_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new ExportErrorsWorkflowService(ThisAddIn.ErrorStore);
                var result = workflow.ExecuteInteractive();

                if (result.Status == ExportErrorsWorkflowStatus.NoErrors)
                {
                    SafeExecutor.ShowInfo("Экспорт қилинадиган хатолар йўқ.");
                    return;
                }
                if (result.Status == ExportErrorsWorkflowStatus.NoUnresolvedErrors)
                {
                    SafeExecutor.ShowInfo("Тузатилмаган хатолар йўқ.");
                    return;
                }
                if (result.Status != ExportErrorsWorkflowStatus.Completed)
                    return;

                if (result.FileKind == ExportErrorsFileKind.Xlsx)
                {
                    ToastNotification.Success("Экспорт тугади", $"{result.ExportedCount} та хато Excelга сақланди.");
                    return;
                }
                if (result.FileKind == ExportErrorsFileKind.Xls)
                {
                    ToastNotification.Success("Экспорт тугади", $"{result.ExportedCount} та хато XLS файлга сақланди.");
                    return;
                }
                ToastNotification.Success("Экспорт тугади", $"{result.ExportedCount} та хато сақланди.");
            }, "Экспорт");
        }

        // =================================================================
        //  GROUP 7: МАХСУС БЕЛГИЛАР (Special Characters)
        // =================================================================

        /// <summary>btnSpecialChar1 — Insert left single quote (tutuq belgi)</summary>
        private void btnSpecialChar1_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new SpecialCharactersWorkflowService(new UzbekApostropheService());
                workflow.InsertCharacter("‘");
            }, "Махсус белги");
        }

        /// <summary>btnSpecialChar2 — Insert right single quote (tutuq belgi)</summary>
        private void btnSpecialChar2_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new SpecialCharactersWorkflowService(new UzbekApostropheService());
                workflow.InsertCharacter("’");
            }, "Махсус белги");
        }

        /// <summary>btnSpecialCharAll — Correct apostrophe/tutuq marks in whole document</summary>
        private void btnSpecialCharAll_Click(object sender, RibbonControlEventArgs e)
        {
            SafeExecutor.Execute(() =>
            {
                var workflow = new SpecialCharactersWorkflowService(new UzbekApostropheService());
                var result = workflow.NormalizeWholeDocument(this.btnSpecialCharAll.Image);
                if (result.Status == SpecialCharactersWorkflowStatus.DocumentUnavailable)
                {
                    SafeExecutor.ShowInfo("Фаол ҳужжат топилмади.");
                    return;
                }
                if (result.Status == SpecialCharactersWorkflowStatus.Cancelled ||
                    result.Status == SpecialCharactersWorkflowStatus.NoOp)
                    return;
                if (result.CorrectedCount > 0)
                {
                    ToastNotification.Success("Тузатилди", $"{result.CorrectedCount} та тутуқ белгиси янгиланди.");
                }
                else
                {
                    ToastNotification.ShowInfo("Тутуқ белгиси", "Тузатиш керак бўлган белги топилмади.");
                }
            }, "Махсус белгилар");
        }

        // =================================================================
        //  HELPER METHODS
        // =================================================================

        /// <summary>
        /// Replace all occurrences of a word in the active document.
        /// </summary>
        private void ReplaceAllInDocument(string findWord, string replaceWord)
        {
            if (!DocumentHelper.IsDocumentOpen()) return;

            var doc = DocumentHelper.ActiveDoc;
            if (doc == null || doc.Content == null) return;

            DocumentHelper.BeginUndoRecord("Барчасини алмаштириш");
            try
            {
                var find = doc.Content.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = findWord;
                find.Replacement.Text = replaceWord;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindContinue;
                find.MatchWholeWord = true;
                find.MatchCase = false;
                find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }
            finally
            {
                DocumentHelper.EndUndoRecord();
            }
        }

        private Image ResolveDefinitionsTitleIcon()
        {
            if (btnDefinitions != null && btnDefinitions.Image != null)
                return btnDefinitions.Image;

            // Thesaurus icon fallback (used by btnDefinitions OfficeImageId).
            var mso = GetOfficeImageMso("Thesaurus", 32, 32);
            if (mso != null) return mso;

            // Stable local fallback when Office icon lookup is unavailable.
            return CreateDefinitionsFallbackIcon(32);
        }

        private Image CreateDefinitionsFallbackIcon(int size)
        {
            if (size < 16) size = 16;

            try
            {
                var bmp = new Bitmap(size, size);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.Clear(Color.Transparent);

                    using (var bg = new SolidBrush(ThemeManager.Primary))
                    {
                        g.FillEllipse(bg, 0, 0, size - 1, size - 1);
                    }

                    using (var border = new Pen(Color.FromArgb(80, 255, 255, 255), 1f))
                    {
                        g.DrawEllipse(border, 0, 0, size - 1, size - 1);
                    }

                    using (var font = new Font("Segoe UI", size * 0.52f, FontStyle.Bold, GraphicsUnit.Pixel))
                    using (var textBrush = new SolidBrush(Color.White))
                    {
                        var sf = new StringFormat
                        {
                            Alignment = StringAlignment.Center,
                            LineAlignment = StringAlignment.Center
                        };
                        g.DrawString("i", font, textBrush, new RectangleF(0, 0, size, size - 1), sf);
                    }
                }
                return bmp;
            }
            catch
            {
                return SystemIcons.Application.ToBitmap();
            }
        }

        private Image GetOfficeImageMso(string imageMso, int width, int height)
        {
            object pictureDisp = null;
            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null || app.CommandBars == null)
                    return null;

                pictureDisp = app.CommandBars.GetImageMso(imageMso, width, height);
                return OfficeImageConverter.FromIPictureDisp(pictureDisp);
            }
            catch
            {
                return null;
            }
            finally
            {
                if (pictureDisp != null && Marshal.IsComObject(pictureDisp))
                {
                    try { Marshal.ReleaseComObject(pictureDisp); } catch { }
                }
            }
        }

        private sealed class OfficeImageConverter : AxHost
        {
            private OfficeImageConverter() : base(string.Empty) { }

            public static Image FromIPictureDisp(object pictureDisp)
            {
                if (pictureDisp == null) return null;
                try { return (Image)GetPictureFromIPictureDisp(pictureDisp); }
                catch { return null; }
            }
        }

        // ─── Маълумот ────────────────────────────────────────────
        private void btnAppInfo_Click(object sender, RibbonControlEventArgs e)
        {
            AppInfoForm.ShowInfo();
        }

        // =================================================================
        //  HOTKEY REGISTRATION
        // =================================================================

        private void RegisterHotkeys()
        {
            try
            {
                var CA = HotkeyManager.Modifiers.Ctrl |
                         HotkeyManager.Modifiers.Alt;
                var CAS = HotkeyManager.Modifiers.Ctrl |
                          HotkeyManager.Modifiers.Alt |
                          HotkeyManager.Modifiers.Shift;

                HotkeyManager.Register(new[]
                {
                    // ── Текшириш ────────────────────────────────────────
                    new HotkeyManager.HotkeyDef(CA,        Keys.Q,      TriggerCheckSpelling,    "Ctrl+Alt+Q"),
                    new HotkeyManager.HotkeyDef(CA,        Keys.E,      TriggerViewErrors,       "Ctrl+Alt+E"),
                    new HotkeyManager.HotkeyDef(CA,        Keys.W,      TriggerErrorList,        "Ctrl+Alt+W"),

                    // ── Тузатиш ─────────────────────────────────────────
                    new HotkeyManager.HotkeyDef(CA,        Keys.J,      TriggerSuggestions,      "Ctrl+Alt+J"),
                    new HotkeyManager.HotkeyDef(CA,        Keys.H,      TriggerReplaceAll,       "Ctrl+Alt+H"),
                    new HotkeyManager.HotkeyDef(CA,        Keys.A,      TriggerToggleAutoCorrect,"Ctrl+Alt+A"),

                    // ── Луғат ───────────────────────────────────────────
                    new HotkeyManager.HotkeyDef(CA,        Keys.B,      TriggerAddToDict,        "Ctrl+Alt+B"),
                    new HotkeyManager.HotkeyDef(CA,        Keys.G,      TriggerEditDictionary,   "Ctrl+Alt+G"),
                    new HotkeyManager.HotkeyDef(CA,        Keys.Y,      TriggerAddNewWords,      "Ctrl+Alt+Y"),

                    // ── Транслитерация ──────────────────────────────────
                    new HotkeyManager.HotkeyDef(CAS,       Keys.C,      TriggerLatinToCyrillic,  "Ctrl+Alt+Shift+C"),
                    new HotkeyManager.HotkeyDef(CAS,       Keys.K,      TriggerCyrillicToLatin,  "Ctrl+Alt+Shift+K"),
                    new HotkeyManager.HotkeyDef(CA,        Keys.X,      TriggerTransExceptions,  "Ctrl+Alt+X"),

                    // ── Воситалар ───────────────────────────────────────
                    new HotkeyManager.HotkeyDef(CAS,       Keys.B,      TriggerCleanupSpaces,    "Ctrl+Alt+Shift+B"),
                    new HotkeyManager.HotkeyDef(CA,        Keys.U,      TriggerSwitchScript,     "Ctrl+Alt+U"),
                    new HotkeyManager.HotkeyDef(CAS,       Keys.N,      TriggerSetFontTNR,       "Ctrl+Alt+Shift+N"),

                    // ── Изоҳлар ─────────────────────────────────────────
                    new HotkeyManager.HotkeyDef(CAS,       Keys.Z,      TriggerDefinitions,      "Ctrl+Alt+Shift+Z"),
                    new HotkeyManager.HotkeyDef(CAS,       Keys.O,      TriggerExportErrors,     "Ctrl+Alt+Shift+O"),

                    // ── Махсус белгилар ──────────────────────────────────
                    new HotkeyManager.HotkeyDef(CAS,       Keys.F,      TriggerSpecialCharAll,   "Ctrl+Alt+Shift+F"),
                    new HotkeyManager.HotkeyDef(CAS,       Keys.D1,     TriggerSpecialChar1,     "Ctrl+Alt+Shift+1"),
                    new HotkeyManager.HotkeyDef(CAS,       Keys.D2,     TriggerSpecialChar2,     "Ctrl+Alt+Shift+2"),
                });
            }
            catch (Exception ex)
            {
                Logger.Error("RegisterHotkeys failed", ex);
            }
        }

        // =================================================================
        //  INTERNAL HOTKEY TRIGGERS — called by HotkeyManager
        // =================================================================

        internal void TriggerCheckSpelling()    => btnCheckSpelling_Click(null, null);
        internal void TriggerViewErrors()       => btnViewErrors_Click(null, null);
        internal void TriggerErrorList()        => btnErrorList_Click(null, null);
        internal void TriggerSuggestions()      => btnSuggestions_Click(null, null);
        internal void TriggerReplaceAll()       => btnReplaceAll_Click(null, null);
        internal void TriggerToggleAutoCorrect()
        {
            toggleAutoCorrect.Checked = !toggleAutoCorrect.Checked;
            toggleAutoCorrect_Click(null, null);
        }
        internal void TriggerAddToDict()        => btnAddToDict_Click(null, null);
        internal void TriggerEditDictionary()   => btnEditDictionary_Click(null, null);
        internal void TriggerAddNewWords()      => btnAddNewWords_Click(null, null);
        internal void TriggerLatinToCyrillic()  => btnFromLatinToCyrillic_Click(null, null);
        internal void TriggerCyrillicToLatin()  => btnFromCyrillicToLatin_Click(null, null);
        internal void TriggerTransExceptions()  => btnTransExceptions_Click(null, null);
        internal void TriggerCleanupSpaces()    => btnCleanupSpaces_Click(null, null);
        internal void TriggerSwitchScript()     => btnSwitchScript_Click(null, null);
        internal void TriggerSetFontTNR()       => btnSetFontTNR_Click(null, null);
        internal void TriggerDefinitions()      => btnDefinitions_Click(null, null);
        internal void TriggerExportErrors()     => btnExportErrors_Click(null, null);
        internal void TriggerSpecialChar1()     => btnSpecialChar1_Click(null, null);
        internal void TriggerSpecialChar2()     => btnSpecialChar2_Click(null, null);
        internal void TriggerSpecialCharAll()   => btnSpecialCharAll_Click(null, null);

    }
}
