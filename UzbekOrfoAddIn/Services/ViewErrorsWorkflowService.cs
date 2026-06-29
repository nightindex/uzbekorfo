using System;
using System.Drawing;
using System.Linq;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Encapsulates the "View Errors" dialog orchestration used by ribbon actions.
    /// </summary>
    public sealed class ViewErrorsWorkflowService
    {
        private readonly ErrorStore _errorStore;
        private readonly Func<string, AddWordResult> _addWordToDictionary;

        public ViewErrorsWorkflowService(
            ErrorStore errorStore,
            Func<string, AddWordResult> addWordToDictionary)
        {
            _errorStore = errorStore;
            _addWordToDictionary = addWordToDictionary;
        }

        public void Execute(Image titleIcon)
        {
            if (_errorStore == null || _errorStore.Count == 0)
            {
                SafeExecutor.ShowInfo("Аввал имло ёки грамматика текширувини ишга туширинг.");
                return;
            }

            var spellingErrors = _errorStore.GetSpellingErrors().Where(e => !e.IsResolved).ToList();
            var grammarErrors = _errorStore.GetGrammarErrors().Where(e => !e.IsResolved).ToList();

            if (spellingErrors.Count == 0 && grammarErrors.Count == 0)
            {
                SafeExecutor.ShowInfo("Хатолар топилмади.");
                return;
            }

            var form = new ViewErrorsForm(
                spellingErrors,
                grammarErrors,
                onReplace: (error, suggestion) =>
                {
                    try
                    {
                        if (error?.Range != null)
                            DocumentHelper.ReplaceRangeText(error.Range, suggestion);
                        _errorStore.RemoveError(error);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Алмаштиришда хато", ex);
                    }
                },
                onIgnore: (error) =>
                {
                    try { _errorStore.RemoveError(error); }
                    catch (Exception ex) { Logger.Error("Ўтказишда хато", ex); }
                },
                onAddToDict: (error) =>
                {
                    try
                    {
                        if (_addWordToDictionary == null)
                            return AddWordResult.Invalid;

                        return _addWordToDictionary(error?.Word?.Trim());
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Луғатга қўшишда хато", ex);
                        return AddWordResult.Invalid;
                    }
                }
            );

            try { form.TitleIcon = titleIcon; }
            catch { }

            form.Show();
        }
    }
}
