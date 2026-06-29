using System;
using System.Drawing;
using UzbekOrfoAddIn.Core;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates dictionary editor dialog setup and execution.
    /// The form is created once and reused on subsequent opens to avoid
    /// the cost of rebuilding ~80+ controls every time.
    /// </summary>
    public sealed class EditDictionaryWorkflowService
    {
        private readonly DictionaryService _dictionaryService;
        private readonly IExplanationProvider _explanationProvider;

        /// <summary>
        /// Cached form instance.  Created on first Execute, reused afterwards.
        /// ShowDialog() hides (not disposes) on close, so the same instance
        /// can be shown again without rebuilding the control tree.
        /// </summary>
        private static EditDictionaryForm _cachedForm;

        public EditDictionaryWorkflowService(
            DictionaryService dictionaryService,
            IExplanationProvider explanationProvider)
        {
            _dictionaryService = dictionaryService;
            _explanationProvider = explanationProvider;
        }

        /// <summary>
        /// Pre-creates the form shell (no data) on the UI thread so the first
        /// user click opens the editor instantly. Words load asynchronously
        /// after the form is shown. Must be called on the UI thread.
        /// </summary>
        public static void PreWarmForm(
            DictionaryService dictionaryService,
            IExplanationProvider explanationProvider)
        {
            if (_cachedForm != null && !_cachedForm.IsDisposed) return;
            if (dictionaryService == null) return;

            try
            {
                // Build only the UI shell — no data. Words will load via
                // QueueInitialWordLoad after Shown fires.
                _cachedForm = new EditDictionaryForm(
                    getWords: () => dictionaryService.GetAllWords(),
                    onAddWord: word => dictionaryService.AddWord(word),
                    onRemoveWord: word =>
                    {
                        dictionaryService.RemoveWord(word);
                        dictionaryService.Save();
                    },
                    isMainWord: word => dictionaryService.IsMainDictionaryWord(word),
                    onCacheBuilt: c => dictionaryService.SetEditorCache(c),
                    explanationProvider: explanationProvider,
                    cachedData: null);

                Logger.Info("Editor form pre-warmed (shell only)");
            }
            catch (Exception ex)
            {
                Logger.Warn($"Editor form pre-warm failed (non-fatal): {ex.Message}");
            }
        }

        public EditDictionaryWorkflowResult Execute(Image titleIcon)
        {
            if (_dictionaryService == null)
                return EditDictionaryWorkflowResult.ServiceUnavailable();

            var cached = _dictionaryService.GetEditorCache();

            if (_cachedForm != null && !_cachedForm.IsDisposed)
            {
                // Reuse existing form — just refresh data (instant if cache hit).
                _cachedForm.Reinitialize(cached);
                try { _cachedForm.TitleIcon = titleIcon; } catch { }
            }
            else
            {
                // First open — build the form from scratch.
                _cachedForm = new EditDictionaryForm(
                    getWords: () => _dictionaryService.GetAllWords(),
                    onAddWord: word =>
                    {
                        _dictionaryService.AddWord(word);
                    },
                    onRemoveWord: word =>
                    {
                        _dictionaryService.RemoveWord(word);
                        _dictionaryService.Save();
                    },
                    isMainWord: word => _dictionaryService.IsMainDictionaryWord(word),
                    onCacheBuilt: c => _dictionaryService.SetEditorCache(c),
                    explanationProvider: _explanationProvider,
                    cachedData: cached);

                try { _cachedForm.TitleIcon = titleIcon; } catch { }
            }

            _cachedForm.ShowDialog();

            return EditDictionaryWorkflowResult.Completed();
        }
    }

    public sealed class EditDictionaryWorkflowResult
    {
        public EditDictionaryWorkflowStatus Status { get; private set; }

        public static EditDictionaryWorkflowResult ServiceUnavailable() =>
            new EditDictionaryWorkflowResult
            {
                Status = EditDictionaryWorkflowStatus.ServiceUnavailable
            };

        public static EditDictionaryWorkflowResult Completed() =>
            new EditDictionaryWorkflowResult
            {
                Status = EditDictionaryWorkflowStatus.Completed
            };
    }

    public enum EditDictionaryWorkflowStatus
    {
        ServiceUnavailable = 0,
        Completed = 1
    }
}
