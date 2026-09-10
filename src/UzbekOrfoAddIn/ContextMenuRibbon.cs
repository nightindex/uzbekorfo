using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn
{
    /// <summary>
    /// Ribbon XML context-menu provider.
    /// Implements native-like "Имло" dynamic submenu for misspelled words.
    /// </summary>
    [ComVisible(true)]
    public class ContextMenuRibbon : Office.IRibbonExtensibility
    {
        private static Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("UzbekOrfoAddIn.ContextMenuRibbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public static void InvalidateContextControls()
        {
            try
            {
                _ribbon?.InvalidateControl("UzbekOrfo_SpellingMenu");
                _ribbon?.InvalidateControl("UzbekOrfo_AddToDict");
                _ribbon?.InvalidateControl("UzbekOrfo_Suggestions");
            }
            catch { }
        }

        public bool GetSpellingMenuVisible(Office.IRibbonControl control)
        {
            try { return ThisAddIn.CtxMenuIsMisspelled; }
            catch { return false; }
        }

        public string GetSpellingMenuContent(Office.IRibbonControl control)
        {
            try
            {
                var suggestions = ThisAddIn.CtxMenuSuggestions ?? new List<string>();
                string rawWord = ThisAddIn.CtxMenuRawWord;

                var sb = new StringBuilder();
                sb.Append("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");

                int i = 0;
                foreach (var suggestion in suggestions)
                {
                    if (string.IsNullOrWhiteSpace(suggestion)) continue;
                    string esc = XmlEscape(suggestion);
                    sb.AppendFormat(
                        "<button id=\"UzbekOrfo_sug_{0}\" label=\"{1}\" tag=\"{1}\" onAction=\"OnSuggestionClick\" />",
                        i++, esc);
                }

                if (i == 0)
                {
                    sb.Append("<button id=\"UzbekOrfo_noSug\" label=\"Вариант топилмади\" enabled=\"false\" />");
                }

                sb.Append("<menuSeparator id=\"UzbekOrfo_sepDict\" />");
                string addLabel = !string.IsNullOrWhiteSpace(rawWord) && rawWord.Length >= 2
                    ? "Луғатга қўшиш:  \"" + XmlEscape(rawWord) + "\""
                    : "Луғатга қўшиш";

                sb.AppendFormat(
                    "<button id=\"UzbekOrfo_addInMenu\" label=\"{0}\" imageMso=\"AddToDictionary\" onAction=\"OnAddToDict\" />",
                    addLabel);

                sb.Append("</menu>");
                return sb.ToString();
            }
            catch (Exception ex)
            {
                Logger.Error("GetSpellingMenuContent error", ex);
                return "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"></menu>";
            }
        }

        public bool GetAddToDictVisible(Office.IRibbonControl control)
        {
            try
            {
                return !ThisAddIn.CtxMenuIsMisspelled &&
                       !string.IsNullOrWhiteSpace(ThisAddIn.CtxMenuRawWord) &&
                       ThisAddIn.CtxMenuRawWord.Length >= 2;
            }
            catch
            {
                return false;
            }
        }

        public string GetAddToDictLabel(Office.IRibbonControl control)
        {
            try
            {
                string word = ThisAddIn.CtxMenuRawWord;
                if (!string.IsNullOrWhiteSpace(word) && word.Length >= 2)
                    return $"Луғатга қўшиш:  \"{word}\"";
            }
            catch { }

            return "Луғатга қўшиш";
        }

        public void OnSuggestionClick(Office.IRibbonControl control)
        {
            try
            {
                string suggestion = control?.Tag;
                if (string.IsNullOrWhiteSpace(suggestion)) return;
                ThisAddIn.ReplaceContextTargetWord(suggestion);
            }
            catch (Exception ex)
            {
                Logger.Error("Context menu suggestion click error", ex);
            }
        }

        public void OnAddToDict(Office.IRibbonControl control)
        {
            try
            {
                string word = ThisAddIn.CtxMenuRawWord;
                if (string.IsNullOrWhiteSpace(word))
                    word = DocumentHelper.GetSelectedWord();
                if (string.IsNullOrWhiteSpace(word))
                {
                    ToastNotification.ShowWarning("Аввал сўзни танланг.");
                    return;
                }

                ThisAddIn.AddWordToUserDictionary(word.Trim(), clearSelectionUnderline: true);
            }
            catch (Exception ex)
            {
                Logger.Error("Context menu AddToDict error", ex);
            }
        }

        /// <summary>
        /// Returns true when a word is selected — enables the Вариантлар context button.
        /// </summary>
        public bool GetSuggestionsVisible(Office.IRibbonControl control)
        {
            try
            {
                string word = ThisAddIn.CtxMenuRawWord;
                return !string.IsNullOrWhiteSpace(word) && word.Length >= 2;
            }
            catch { return false; }
        }

        /// <summary>
        /// Context menu "Вариантлар" — opens the full suggestions dialog.
        /// </summary>
        public void OnSuggestionsClick(Office.IRibbonControl control)
        {
            try
            {
                if (!DocumentHelper.IsDocumentOpen()) return;

                var workflow = new Services.SuggestionsWorkflowService(
                    ThisAddIn.SpellingEngine,
                    ThisAddIn.GrammarEngine,
                    Globals.ThisAddIn?.Application,
                    word => ThisAddIn.AddWordToUserDictionary(word),
                    (fromWord, toWord) => DocumentHelper.ReplaceAllInDocument(fromWord, toWord));

                workflow.Execute(null);
            }
            catch (Exception ex)
            {
                Logger.Error("Context menu Suggestions error", ex);
            }
        }

        private static string XmlEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            for (int i = 0; i < resourceNames.Length; i++)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var reader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        return reader?.ReadToEnd();
                    }
                }
            }

            return null;
        }
    }
}
