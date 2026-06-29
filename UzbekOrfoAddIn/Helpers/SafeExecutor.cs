using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;
using Word = Microsoft.Office.Interop.Word;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Provides a safe execution wrapper and common UI helper methods.
    /// All ribbon button handlers should use SafeExecute to ensure errors are caught.
    /// </summary>
    public static class SafeExecutor
    {
        /// <summary>
        /// Wraps an action in try-catch with user-friendly error messages.
        /// Use this in every ribbon button handler that requires an open document.
        /// </summary>
        public static void Execute(Action action, string operationName = "Амал")
        {
            Execute(action, operationName, requiresDocument: true);
        }

        /// <summary>
        /// Wraps an action in try-catch with user-friendly error messages.
        /// Set <paramref name="requiresDocument"/> to false for operations that
        /// do not need an open document (settings, dictionary management, about, etc.).
        /// </summary>
        public static void Execute(Action action, string operationName, bool requiresDocument)
        {
            try
            {
                if (!ThisAddIn.IsInitialized)
                {
                    ShowWarning("Қўшимча тўлиқ юкланмаган. Илтимос, Word-ни қайта ишга туширинг.");
                    return;
                }

                if (requiresDocument && !DocumentHelper.IsDocumentOpen())
                {
                    ShowWarning("Илтимос, аввал ҳужжат очинг.");
                    return;
                }

                action();
            }
            catch (COMException ex)
            {
                Logger.Error(operationName, ex);
                ShowError($"{operationName}: Word билан алоқада хатолик юз берди.\n\n{ex.Message}");
            }
            catch (System.IO.IOException ex)
            {
                Logger.Error(operationName, ex);
                ShowError($"{operationName}: Файл билан ишлашда хатолик.\n\n{ex.Message}");
            }
            catch (Exception ex)
            {
                Logger.Error(operationName, ex);
                ShowError($"{operationName}: Кутилмаган хатолик юз берди.\n\n{ex.Message}");
            }
        }

        /// <summary>
        /// Shows a modern error message dialog.
        /// </summary>
        public static void ShowError(string message)
        {
            ModernMessageBox.Error(message);
        }

        /// <summary>
        /// Shows a modern warning message dialog.
        /// </summary>
        public static void ShowWarning(string message)
        {
            ModernMessageBox.Warning(message);
        }

        /// <summary>
        /// Shows a modern informational message dialog.
        /// </summary>
        public static void ShowInfo(string message)
        {
            ModernMessageBox.Show(message);
        }

        /// <summary>
        /// Shows a modern Yes/No confirmation dialog. Returns true if user clicks Yes.
        /// </summary>
        public static bool Confirm(string message)
        {
            return ModernMessageBox.Confirm(message);
        }
    }
}
