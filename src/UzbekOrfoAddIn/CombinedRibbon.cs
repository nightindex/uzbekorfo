using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Office = Microsoft.Office.Core;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn
{
    /// <summary>
    /// Proxy that wraps the VSTO internal RibbonManager (which handles the
    /// Ribbon Designer tab) and injects native Fluent UI context-menu
    /// customizations into its XML output.
    ///
    /// Architecture:
    ///   1. ThisAddIn.CreateRibbonExtensibilityObject() returns this proxy,
    ///      passing the base VSTO RibbonManager as the inner object.
    ///   2. GetCustomUI() merges the RibbonManager's ribbon-tab XML with
    ///      our &lt;contextMenus&gt; fragment.
    ///   3. IReflect routes COM IDispatch callbacks:
    ///        - Context-menu callbacks  ->  handled here
    ///        - Ribbon-tab callbacks    ->  forwarded to inner RibbonManager
    ///   4. Result: Designer ribbon + native context menus coexist.
    /// </summary>
    [ComVisible(true)]
    public sealed class CombinedRibbon : Office.IRibbonExtensibility, IReflect
    {
        private readonly object _inner;
        private readonly Type _innerType;
        private string _innerOnLoadName;

        private static Office.IRibbonUI _ribbon;

        public CombinedRibbon(object innerRibbonManager)
        {
            _inner = innerRibbonManager;
            _innerType = innerRibbonManager?.GetType();
        }

        // ==================================================================
        //  IRibbonExtensibility
        // ==================================================================

        public string GetCustomUI(string ribbonID)
        {
            string baseXml = string.Empty;
            try
            {
                var ext = _inner as Office.IRibbonExtensibility;
                if (ext != null)
                    baseXml = ext.GetCustomUI(ribbonID) ?? string.Empty;
            }
            catch (Exception ex)
            {
                Logger.Error("CombinedRibbon: GetCustomUI from inner failed", ex);
            }

            if (string.IsNullOrEmpty(baseXml))
                return StandaloneContextMenuXml();

            return MergeXml(baseXml);
        }

        // ==================================================================
        //  XML Merging
        // ==================================================================

        private string MergeXml(string baseXml)
        {
            // contextMenus require the 2009 namespace; upgrade if needed.
            if (baseXml.Contains("http://schemas.microsoft.com/office/2006/01/customui"))
            {
                baseXml = baseXml.Replace(
                    "http://schemas.microsoft.com/office/2006/01/customui",
                    "http://schemas.microsoft.com/office/2009/07/customui");
            }

            // Intercept onLoad so we can forward to both inner + self.
            var onLoadMatch = Regex.Match(baseXml, @"onLoad\s*=\s*""([^""]+)""");
            if (onLoadMatch.Success)
            {
                _innerOnLoadName = onLoadMatch.Groups[1].Value;
                baseXml = baseXml.Replace(onLoadMatch.Value, "onLoad=\"CombinedOnLoad\"");
            }
            else
            {
                baseXml = Regex.Replace(baseXml, @"<customUI\b", "<customUI onLoad=\"CombinedOnLoad\"");
            }

            // Inject <contextMenus> before </customUI>
            int closeIdx = baseXml.LastIndexOf("</customUI>", StringComparison.OrdinalIgnoreCase);
            if (closeIdx >= 0)
            {
                return baseXml.Substring(0, closeIdx)
                     + ContextMenusFragment()
                     + "\n"
                     + baseXml.Substring(closeIdx);
            }

            return baseXml;
        }

        /// <summary>Context-menu XML fragment injected into the base ribbon XML.</summary>
        private static string ContextMenusFragment()
        {
            // label / imageMso use standard Office Fluent UI identifiers.
            // Cyrillic label uses Unicode escapes because the .cs file
            // encoding can vary across developer machines.
            return
                "  <contextMenus>\n" +
                "    <contextMenu idMso=\"ContextMenuText\">\n" +
                "      <dynamicMenu id=\"UzbekOrfo_SpellingMenu\"\n" +
                "                   getContent=\"GetSpellingMenuContent\"\n" +
                "                   getVisible=\"GetSpellingMenuVisible\"\n" +
                "                   label=\"\u0418\u043c\u043b\u043e\"\n" +
                "                   imageMso=\"SpellingAndGrammar\"\n" +
                "                   insertBeforeMso=\"Cut\" />\n" +
                "      <button id=\"UzbekOrfo_AddToDict\"\n" +
                "              getLabel=\"GetAddToDictLabel\"\n" +
                "              getVisible=\"GetAddToDictVisible\"\n" +
                "              imageMso=\"AddToDictionary\"\n" +
                "              onAction=\"OnAddToDict\"\n" +
                "              insertBeforeMso=\"Cut\" />\n" +
                "      <menuSeparator id=\"UzbekOrfo_SepSuggestions\"\n" +
                "                     insertBeforeMso=\"Cut\" />\n" +
                "      <button id=\"UzbekOrfo_Suggestions\"\n" +
                "              label=\"\u0412\u0430\u0440\u0438\u0430\u043d\u0442\u043b\u0430\u0440\"\n" + // Вариантлар
                "              imageMso=\"AutoCorrect\"\n" +
                "              onAction=\"OnSuggestionsClick\"\n" +
                "              getVisible=\"GetSuggestionsVisible\"\n" +
                "              insertBeforeMso=\"Cut\" />\n" +
                "    </contextMenu>\n" +
                "  </contextMenus>\n";
        }

        /// <summary>Fallback: full XML when the inner manager returns nothing.</summary>
        private static string StandaloneContextMenuXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                   "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"\n" +
                   "          onLoad=\"CombinedOnLoad\">\n" +
                   ContextMenusFragment() +
                   "</customUI>";
        }

        // ==================================================================
        //  Combined onLoad
        // ==================================================================

        public void CombinedOnLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;

            // Forward to inner RibbonManager's original onLoad handler
            // so the Designer ribbon tab initializes correctly.
            if (!string.IsNullOrEmpty(_innerOnLoadName) && _inner != null)
            {
                try
                {
                    _innerType.InvokeMember(
                        _innerOnLoadName,
                        BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                        null, _inner, new object[] { ribbonUI });
                }
                catch (Exception ex)
                {
                    Logger.Error("CombinedOnLoad: forwarding to inner RibbonManager failed", ex);
                }
            }
        }

        /// <summary>
        /// Force Office to re-query the context-menu controls.
        /// Called from ThisAddIn.App_WindowBeforeRightClick after
        /// pre-computing spelling data.
        /// </summary>
        public static void InvalidateContextControls()
        {
            try
            {
                _ribbon?.InvalidateControl("UzbekOrfo_SpellingMenu");
                _ribbon?.InvalidateControl("UzbekOrfo_AddToDict");
                _ribbon?.InvalidateControl("UzbekOrfo_SepSuggestions");
                _ribbon?.InvalidateControl("UzbekOrfo_Suggestions");
            }
            catch { }
        }

        // ==================================================================
        //  Context-Menu Callbacks
        // ==================================================================

        /// <summary>Visible only when the right-clicked word is misspelled.</summary>
        public bool GetSpellingMenuVisible(Office.IRibbonControl control)
        {
            try { return ThisAddIn.CtxMenuIsMisspelled; }
            catch { return false; }
        }

        /// <summary>
        /// Builds the submenu content XML: suggestion buttons + add-to-dictionary.
        /// </summary>
        public string GetSpellingMenuContent(Office.IRibbonControl control)
        {
            try
            {
                var suggestions = ThisAddIn.CtxMenuSuggestions ?? new List<string>();
                string rawWord = ThisAddIn.CtxMenuRawWord;

                var sb = new StringBuilder();
                sb.Append("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");

                int count = 0;
                foreach (var sug in suggestions)
                {
                    if (string.IsNullOrWhiteSpace(sug)) continue;
                    string esc = XmlEscape(sug);
                    sb.AppendFormat(
                        "<button id=\"UzbekOrfo_sug_{0}\" label=\"{1}\" tag=\"{1}\" onAction=\"OnSuggestionClick\" />",
                        count++, esc);
                }

                if (count == 0)
                {
                    // "Вариант топилмади"
                    sb.Append("<button id=\"UzbekOrfo_noSug\" " +
                              "label=\"\u0412\u0430\u0440\u0438\u0430\u043d\u0442 \u0442\u043e\u043f\u0438\u043b\u043c\u0430\u0434\u0438\" " +
                              "enabled=\"false\" />");
                }

                sb.Append("<menuSeparator id=\"UzbekOrfo_sepDict\" />");

                // "Луғатга қўшиш"
                string addLabel;
                if (!string.IsNullOrWhiteSpace(rawWord) && rawWord.Length >= 2)
                    addLabel = "\u041b\u0443\u0493\u0430\u0442\u0433\u0430 \u049b\u045e\u0448\u0438\u0448:  \"" + XmlEscape(rawWord) + "\"";
                else
                    addLabel = "\u041b\u0443\u0493\u0430\u0442\u0433\u0430 \u049b\u045e\u0448\u0438\u0448";

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

        /// <summary>
        /// Standalone "Add to Dictionary" button visible when word is NOT misspelled.
        /// </summary>
        public bool GetAddToDictVisible(Office.IRibbonControl control)
        {
            try
            {
                return !ThisAddIn.CtxMenuIsMisspelled &&
                       !string.IsNullOrWhiteSpace(ThisAddIn.CtxMenuRawWord) &&
                       ThisAddIn.CtxMenuRawWord.Length >= 2;
            }
            catch { return false; }
        }

        /// <summary>Dynamic label showing the word under the cursor.</summary>
        public string GetAddToDictLabel(Office.IRibbonControl control)
        {
            try
            {
                string word = ThisAddIn.CtxMenuRawWord;
                if (!string.IsNullOrWhiteSpace(word) && word.Length >= 2)
                    return "\u041b\u0443\u0493\u0430\u0442\u0433\u0430 \u049b\u045e\u0448\u0438\u0448:  \"" + word + "\"";
            }
            catch { }
            return "\u041b\u0443\u0493\u0430\u0442\u0433\u0430 \u049b\u045e\u0448\u0438\u0448";
        }

        /// <summary>Returns true when a word is selected — enables the Вариантлар context button.</summary>
        public bool GetSuggestionsVisible(Office.IRibbonControl control)
        {
            try
            {
                string word = ThisAddIn.CtxMenuRawWord;
                return !string.IsNullOrWhiteSpace(word) && word.Length >= 2;
            }
            catch { return false; }
        }

        /// <summary>Context menu "Вариантлар" — opens the full suggestions dialog.</summary>
        public void OnSuggestionsClick(Office.IRibbonControl control)
        {
            try
            {
                if (!DocumentHelper.IsDocumentOpen()) return;

                System.Drawing.Image icon = null;
                try { icon = Globals.Ribbons?.UzbekOrfoRibbon?.btnSuggestions?.Image; } catch { }

                var workflow = new Services.SuggestionsWorkflowService(
                    ThisAddIn.SpellingEngine,
                    ThisAddIn.GrammarEngine,
                    Globals.ThisAddIn?.Application,
                    word => ThisAddIn.AddWordToUserDictionary(word),
                    (fromWord, toWord) => DocumentHelper.ReplaceAllInDocument(fromWord, toWord));

                workflow.Execute(icon);
            }
            catch (Exception ex)
            {
                Logger.Error("Context menu Suggestions error", ex);
            }
        }

        /// <summary>User clicked a suggestion in the spelling submenu.</summary>
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
                Logger.Error("Context-menu suggestion click error", ex);
            }
        }

        /// <summary>User clicked "Add to Dictionary".</summary>
        public void OnAddToDict(Office.IRibbonControl control)
        {
            try
            {
                string word = ThisAddIn.CtxMenuRawWord;
                if (string.IsNullOrWhiteSpace(word))
                    word = DocumentHelper.GetSelectedWord();
                if (string.IsNullOrWhiteSpace(word))
                {
                    // "Аввал сўзни танланг."
                    ToastNotification.ShowWarning(
                        "\u0410\u0432\u0432\u0430\u043b \u0441\u045e\u0437\u043d\u0438 \u0442\u0430\u043d\u043b\u0430\u043d\u0433.");
                    return;
                }
                ThisAddIn.AddWordToUserDictionary(word.Trim(), clearSelectionUnderline: true);
            }
            catch (Exception ex)
            {
                Logger.Error("Context-menu AddToDict error", ex);
            }
        }

        // ==================================================================
        //  IReflect  —  Routes COM IDispatch between self and inner manager
        //
        //  When Office calls a callback by name (e.g. "GetSpellingMenuVisible"),
        //  the CLR's CCW uses IReflect to resolve and invoke it.
        //  * Methods declared on CombinedRibbon  →  executed here
        //  * Everything else                     →  forwarded to RibbonManager
        // ==================================================================

        Type IReflect.UnderlyingSystemType
        {
            get { return typeof(CombinedRibbon); }
        }

        object IReflect.InvokeMember(
            string name, BindingFlags invokeAttr, Binder binder,
            object target, object[] args, ParameterModifier[] modifiers,
            CultureInfo culture, string[] namedParameters)
        {
            // 1. Our own declared methods take priority.
            var ownMethod = typeof(CombinedRibbon).GetMethod(
                name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
            if (ownMethod != null)
                return ownMethod.Invoke(this, args);

            // 2. Forward to inner RibbonManager for ribbon-tab callbacks.
            if (_inner != null)
            {
                try
                {
                    return _innerType.InvokeMember(
                        name, invokeAttr, binder, _inner,
                        args, modifiers, culture, namedParameters);
                }
                catch (MissingMethodException) { }
                catch (MissingMemberException) { }
            }

            // 3. Last resort — fall back to our full type (including inherited).
            var fallback = typeof(CombinedRibbon).GetMethod(
                name, BindingFlags.Public | BindingFlags.Instance);
            if (fallback != null)
                return fallback.Invoke(this, args);

            return null;
        }

        MethodInfo IReflect.GetMethod(string name, BindingFlags bindingAttr)
        {
            var m = typeof(CombinedRibbon).GetMethod(name, bindingAttr);
            if (m != null) return m;
            return _innerType?.GetMethod(name, bindingAttr);
        }

        MethodInfo IReflect.GetMethod(
            string name, BindingFlags bindingAttr,
            Binder binder, Type[] types, ParameterModifier[] modifiers)
        {
            var m = typeof(CombinedRibbon).GetMethod(name, bindingAttr, binder, types, modifiers);
            if (m != null) return m;
            return _innerType?.GetMethod(name, bindingAttr, binder, types, modifiers);
        }

        MethodInfo[] IReflect.GetMethods(BindingFlags bindingAttr)
        {
            var own = typeof(CombinedRibbon).GetMethods(bindingAttr);
            var inner = _innerType != null
                ? _innerType.GetMethods(bindingAttr)
                : new MethodInfo[0];
            var combined = new MethodInfo[own.Length + inner.Length];
            own.CopyTo(combined, 0);
            inner.CopyTo(combined, own.Length);
            return combined;
        }

        MemberInfo[] IReflect.GetMember(string name, BindingFlags bindingAttr)
        {
            var own = typeof(CombinedRibbon).GetMember(name, bindingAttr);
            if (own.Length > 0) return own;
            return _innerType != null
                ? _innerType.GetMember(name, bindingAttr)
                : new MemberInfo[0];
        }

        MemberInfo[] IReflect.GetMembers(BindingFlags bindingAttr)
        {
            var own = typeof(CombinedRibbon).GetMembers(bindingAttr);
            var inner = _innerType != null
                ? _innerType.GetMembers(bindingAttr)
                : new MemberInfo[0];
            var combined = new MemberInfo[own.Length + inner.Length];
            own.CopyTo(combined, 0);
            inner.CopyTo(combined, own.Length);
            return combined;
        }

        FieldInfo IReflect.GetField(string name, BindingFlags bindingAttr)
        {
            return typeof(CombinedRibbon).GetField(name, bindingAttr)
                ?? (_innerType != null ? _innerType.GetField(name, bindingAttr) : null);
        }

        FieldInfo[] IReflect.GetFields(BindingFlags bindingAttr)
        {
            return typeof(CombinedRibbon).GetFields(bindingAttr);
        }

        PropertyInfo IReflect.GetProperty(string name, BindingFlags bindingAttr)
        {
            return typeof(CombinedRibbon).GetProperty(name, bindingAttr)
                ?? (_innerType != null ? _innerType.GetProperty(name, bindingAttr) : null);
        }

        PropertyInfo IReflect.GetProperty(
            string name, BindingFlags bindingAttr,
            Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
        {
            var p = typeof(CombinedRibbon).GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
            if (p != null) return p;
            return _innerType != null
                ? _innerType.GetProperty(name, bindingAttr, binder, returnType, types, modifiers)
                : null;
        }

        PropertyInfo[] IReflect.GetProperties(BindingFlags bindingAttr)
        {
            return typeof(CombinedRibbon).GetProperties(bindingAttr);
        }

        // ==================================================================
        //  Helpers
        // ==================================================================

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
    }
}
