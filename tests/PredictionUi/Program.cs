using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Prediction;
using UzbekOrfoAddIn.UI;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        string root = Path.Combine(Path.GetTempPath(), "MatnAI-Ui-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            bool dark = args.Contains("--dark");
            typeof(ThemeManager).GetField("_isDarkTheme", BindingFlags.NonPublic | BindingFlags.Static).SetValue(null, (bool?)dark);
            using ((IDisposable)Activator.CreateInstance(typeof(DpiForm).Assembly.GetType("UzbekOrfoAddIn.UI.DpiLayout+Context"), true))
            {
            var store = new CollectionStore(Path.Combine(root, "models"));
            var builtIns = store.LoadAll();
            Require(builtIns.Count == 2 && builtIns.All(c => c.BuiltIn) &&
                builtIns.Any(c => c.Id == "legal-default") && builtIns.Any(c => c.Id == "legal-latin"),
                "Packaged add-in loads both embedded legal models on .NET Framework");
            string[] active = new string[0]; int changed = 0;
            var control = new MatnAiCollectionsControl(store, () => changed++, () => active, ids => active = ids);
            using (var form = new MatnAiSettingsForm(2, 3, false, (a, b, c) => { }, () => { }, () => { }, control))
            {
                form.Opacity = 0; form.ShowInTaskbar = false; form.Show();
                var tabs = Find<TabControl>(form); tabs.SelectedIndex = 1;
                Application.DoEvents(); WaitUntil(() => !control.IsBusy);
                var initialCollections = Field<CheckedListBox>(control, "_collections");
                Require(initialCollections.Items.Count == 2 && initialCollections.CheckedItems.Count == 2,
                    "Both built-in legal collections appear enabled in settings");
                var summaries = Field<System.Collections.Generic.List<PredictionCollection>>(control, "_items");
                Require(summaries.All(c => c.Sources.Count > 0 && c.Sources.All(s =>
                    s.Sequences.Count == 0 && s.DisplayWords.Count == 0 && s.Passages.Count == 0)),
                    "Settings retains built-in source metadata without phrase payloads");
                foreach (var original in builtIns)
                {
                    int index = summaries.FindIndex(c => c.Id == original.Id);
                    int phraseCount = original.Sources.SelectMany(s => s.Sequences.Keys).Where(k => k.Contains(" ")).Distinct().Count();
                    Require(initialCollections.Items[index].ToString().Contains(phraseCount + " ибора"),
                        "Built-in phrase count survives lightweight projection");
                    Require(summaries[index].Sources.Select(s => s.Path).SequenceEqual(original.Sources.Select(s => s.Path)),
                        "Built-in source paths remain available for display");
                }
                builtIns.Clear();
                var name = Field<TextBox>(control, "_name");
                string input = Path.Combine(root, "documents"); Directory.CreateDirectory(input);
                File.WriteAllText(Path.Combine(input, "first.txt"), "Sud tomonidan qaror qabul qilindi bugun. Sud tomonidan qaror qabul qilindi kecha.");
                Pump(Run(control, "AddPaths", (object)new[] { input }));
                name.Text = "Шартномалар";
                Pump(Run(control, "StartImport"));
                var personal = store.LoadAll().Single(c => !c.BuiltIn);
                Require(personal.Sources.Count == 1 && personal.ImportLocations.ContainsKey(input), "Folder import records its source and refresh location");
                Require(Field<System.Collections.Generic.List<PredictionCollection>>(control, "_items")
                    .Single(c => !c.BuiltIn).Sources.Any(s => s.Sequences.Count > 0),
                    "Editable private collection retains its learned counts");
                Require(active.Length == 0 && changed == 1, "Import does not silently enable a private collection");
                var collections = Field<CheckedListBox>(control, "_collections");
                int personalIndex = Enumerable.Range(0, collections.Items.Count).Single(i => collections.Items[i].ToString().StartsWith("Шартномалар"));
                collections.SelectedIndex = personalIndex; collections.SetItemChecked(personalIndex, true);
                Require(active.SequenceEqual(new[] { personal.Id }), "Checkbox changes current-document selection");
                File.WriteAllText(Path.Combine(input, "second.txt"), "Yangi hujjat. Yangi qonun.");
                Pump(Run(control, "RefreshCollection"));
                Require(store.LoadAll().Single(c => !c.BuiltIn).Sources.Count == 2, "Manual folder refresh discovers newly added files");
                Application.DoEvents();
                Require(tabs.DrawMode == TabDrawMode.OwnerDrawFixed && tabs.TabCount == 3, "Modern navigation retains three native accessible pages");
                for (int page = 0; page < tabs.TabCount; page++)
                {
                    tabs.SelectedIndex = page; Application.DoEvents();
                    Require(tabs.GetTabRect(page).Right <= tabs.ClientSize.Width, "Page navigation fits the settings width");
                    using (var preview = new Bitmap(form.Width, form.Height))
                    {
                        form.DrawToBitmap(preview, form.ClientRectangle);
                        preview.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings-" + (dark ? "dark" : "light") + "-" + page + ".png"));
                    }
                }
                tabs.SelectedIndex = 1; Application.DoEvents();
                Console.WriteLine("LAYOUT: form=" + form.Size + " content=" + form.ContentPanel.Bounds + " tabs=" + tabs.Bounds + " page=" + tabs.SelectedTab.Bounds + " control=" + control.Bounds + " root=" + collections.Parent.Bounds + " collections=" + collections.Bounds);
                Require(collections.Width <= control.ClientSize.Width, "Initial collection list fits the page");
                using (var initial = new Bitmap(form.Width, form.Height))
                { form.DrawToBitmap(initial, form.ClientRectangle); initial.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "collections-initial.png")); }
                foreach (int dpi in new[] { 96, 144, 192, 96 })
                {
                    typeof(DpiForm).GetMethod("ApplyDpi", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form, new object[] { dpi });
                    form.Size = new Size(900, 800); form.PerformLayout(); Application.DoEvents();
                    Require(collections.Height > 20 && Field<ListBox>(control, "_files").Height > 20, "Collection and source lists remain usable at DPI " + dpi);
                    Require(collections.Width <= control.ClientSize.Width, "Collection list fits the page width at DPI " + dpi);
                    Require(tabs.Right <= form.ContentPanel.ClientSize.Width, "Tabs fit content at DPI " + dpi);
                    Console.WriteLine("DPI " + dpi + ": content=" + form.ContentPanel.Bounds + "; tabs=" + tabs.Bounds + "; visible=" + collections.Visible);
                    Require(!form.ActionBar.Parent.RectangleToScreen(form.ActionBar.Parent.ClientRectangle).IntersectsWith(
                        form.ContentPanel.Parent.RectangleToScreen(form.ContentPanel.Parent.ClientRectangle)), "Action and content viewports remain separate");
                }
                using (var screenshot = new Bitmap(form.Width, form.Height))
                { form.DrawToBitmap(screenshot, form.ClientRectangle); screenshot.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "collections.png")); }
                Console.WriteLine("PASS: real collection UI folder import, manual refresh, opt-in selection and DPI layout");
            }
            }
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { Directory.Delete(root, true); }
    }
    private static Task Run(object target, string method, params object[] args)
    {
        Func<Task> action = () => (Task)target.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, args);
        return (Task)target.GetType().GetMethod("Run", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, new object[] { action });
    }
    private static T Field<T>(object target, string name) => (T)target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
    private static T Find<T>(Control root) where T : Control
    {
        if (root is T) return (T)root;
        foreach (Control child in root.Controls) { var found = Find<T>(child); if (found != null) return found; }
        return null;
    }
    private static void Pump(Task task) { WaitUntil(() => task.IsCompleted); task.GetAwaiter().GetResult(); }
    private static void WaitUntil(Func<bool> completed)
    {
        var timeout = Stopwatch.StartNew();
        while (!completed() && timeout.ElapsedMilliseconds < 30000) { Application.DoEvents(); Thread.Sleep(5); }
        Require(completed(), "UI work completed before timeout");
    }
    private static void Require(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
}
