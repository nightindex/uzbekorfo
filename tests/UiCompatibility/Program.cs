using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using UzbekOrfoAddIn.Forms;
using UzbekOrfoAddIn.Models;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

// A Windows/STA smoke test of real controls. No Word instance or user files.
internal static class Program
{
    private sealed class KeyboardToggle : ModernToggle
    {
        public void PressSpace()
        {
            OnKeyDown(new KeyEventArgs(Keys.Space));
            OnKeyDown(new KeyEventArgs(Keys.Space)); // key repeat must not toggle twice
            OnKeyUp(new KeyEventArgs(Keys.Space));
        }
    }

    private static void CheckKeyboardAccessibility()
    {
        using (var toggle = new KeyboardToggle { AccessibleName = "Include context" })
        {
            toggle.PressSpace();
            Require(toggle.IsOn, "Space activates toggle once, even with key repeat");
            Require(toggle.AccessibilityObject.Role == AccessibleRole.CheckButton, "Toggle exposes check-button role");
            Require((toggle.AccessibilityObject.State & AccessibleStates.Checked) != 0, "Toggle exposes checked state");
            toggle.Enabled = false;
            toggle.PressSpace();
            Require(toggle.IsOn, "Disabled toggle ignores keyboard");
            toggle.Enabled = true;
            toggle.AccessibilityObject.DoDefaultAction();
            Require(!toggle.IsOn, "Assistive technology can activate toggle");
        }
        Console.WriteLine("PASS: toggle keyboard and accessibility behavior");
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr window, int id);

    private static void CheckShortcutsDoNotReserveGlobalKeys()
    {
        using (var window = new Control())
        {
            var handle = window.Handle;
            uint key = 0;
            for (uint candidate = (uint)Keys.F13; candidate <= (uint)Keys.F24; candidate++)
                if (RegisterHotKey(handle, 123, 7, candidate))
                {
                    key = candidate;
                    UnregisterHotKey(handle, 123);
                    break;
                }
            Require(key != 0, "An unused test shortcut is available");
            try
            {
                UzbekOrfoAddIn.Helpers.HotkeyManager.Register(new[]
                {
                    new UzbekOrfoAddIn.Helpers.HotkeyManager.HotkeyDef(
                        UzbekOrfoAddIn.Helpers.HotkeyManager.Modifiers.Ctrl |
                        UzbekOrfoAddIn.Helpers.HotkeyManager.Modifiers.Alt |
                        UzbekOrfoAddIn.Helpers.HotkeyManager.Modifiers.Shift,
                        (Keys)key, () => { throw new InvalidOperationException("Unexpected shortcut action"); })
                });
                var hook = typeof(UzbekOrfoAddIn.Helpers.HotkeyManager).GetField("_hook", BindingFlags.Static | BindingFlags.NonPublic);
                Require((IntPtr)hook.GetValue(null) != IntPtr.Zero, "Thread keyboard hook installed");
                Require(RegisterHotKey(handle, 123, 7, key), "Word shortcut leaves system-wide registration available");
            }
            finally
            {
                UnregisterHotKey(handle, 123);
                UzbekOrfoAddIn.Helpers.HotkeyManager.Unregister();
            }
        }
        Console.WriteLine("PASS: shortcuts do not reserve global keys");
    }

    [STAThread]
    private static int Main()
    {
        try
        {
            Application.EnableVisualStyles();
            CheckKeyboardAccessibility();
            CheckShortcutsDoNotReserveGlobalKeys();
            CheckExceptionFonts();
            CheckInitialGridWindow();
            CheckDeferredErrorSuggestions();
            using (new DpiLayout.Context())
            {
                Application.EnableVisualStyles();
                var factories = new Func<ModernForm>[]
                {
                    () => new AddNewWordsForm(), () => new AppInfoForm(), () => new ImportProgressForm(),
                    () => new TranslitExceptionsForm(
                        () => Enumerable.Range(1, 39).Select(i => new TranslitException("Example " + i, "Мисол " + i)).ToList(),
                        item => { throw new InvalidOperationException("Unexpected add"); },
                        item => { throw new InvalidOperationException("Unexpected edit"); },
                        item => { throw new InvalidOperationException("Unexpected delete"); },
                        () => { throw new InvalidOperationException("Unexpected save"); }),
                    () => (ImportResultForm)Activator.CreateInstance(typeof(ImportResultForm),
                        BindingFlags.Instance | BindingFlags.NonPublic, null,
                        new object[] { ImportResultForm.ResultType.Success, 150, 25, 3, 10000, 5 }, null),
                    () => (ModernMessageBox)Activator.CreateInstance(typeof(ModernMessageBox),
                        BindingFlags.Instance | BindingFlags.NonPublic, null,
                        new object[] { string.Join(" ", Enumerable.Repeat("Long message / Ўзбекча матн.", 40)),
                            "Display compatibility", ModernMessageBox.MessageType.Warning, true, "OK", "Cancel", null }, null),
                    () => new ControlTestForm(),
                    () => (ModernMessageBox)Activator.CreateInstance(typeof(ModernMessageBox),
                        BindingFlags.Instance | BindingFlags.NonPublic, null,
                        new object[] { "Матн тўғри!\nИмло ва грамматик хатолар топилмади.",
                            "Матн текшируви", ModernMessageBox.MessageType.Success, false, "Яхши", "", null }, null)
                };
                foreach (int dpi in new[] { 96, 120, 144, 192, 240, 288, 384 })
                {
                    foreach (var factory in factories)
                    using (var form = factory())
                    {
                        Size preferred = form.Size;
                        // Native visibility is necessary for real scrollbar/layout behavior.
                        // Keep test windows fully transparent and off the taskbar.
                        form.Opacity = 0;
                        form.ShowInTaskbar = false;
                        form.Show();
                        Application.DoEvents();
                        var scale = typeof(DpiForm).GetMethod("ApplyDpi", BindingFlags.Instance | BindingFlags.NonPublic);
                        int originalDpi = form.LayoutDpi;
                        int footerHeight = form.ActionBar.Height * 96 / originalDpi;
                        scale.Invoke(form, new object[] { dpi });
                        Require(form.LayoutDpi == dpi, "DPI applied");
                        form.MinimumSize = Size.Empty;
                        form.Size = new Size(ScreenGeometry.Scale(preferred.Width, dpi), ScreenGeometry.Scale(preferred.Height, dpi));
                        form.PerformLayout();
                        var bodyHost = (ScrollableControl)form.ContentPanel.Parent;
                        if (dpi == 120 && form is TranslitExceptionsForm)
                        {
                            using (var screenshot = new Bitmap(form.Width, form.Height))
                            {
                                form.DrawToBitmap(screenshot, form.ClientRectangle);
                                screenshot.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Exceptions-Normal-120.png"));
                            }
                        }
                        if (form is TranslitExceptionsForm)
                        {
                            var grid = Descendants(form).OfType<DataGridView>().Single();
                            Require(grid.Visible && grid.Height >= ScreenGeometry.Scale(100, dpi), "Exceptions grid stays visible");
                            Require(grid.Rows.Count == 39, "Exceptions data remains displayed");
                        }
                        // Windows caps live test windows at the current monitor's tracking size.
                        // Only require no overflow when the requested size was actually granted.
                        if (form.Width >= ScreenGeometry.Scale(preferred.Width, dpi) &&
                            form.Height >= ScreenGeometry.Scale(preferred.Height, dpi))
                            Require(!bodyHost.HorizontalScroll.Visible && !bodyHost.VerticalScroll.Visible,
                                "No outer scrollbars at preferred size: " + form.GetType().Name + " at " + dpi);
                        if (form.Title == "Display compatibility")
                            Require(form.ContentPanel.VerticalScroll.Visible,
                                "Long message retains necessary vertical scrolling");
                        if (form.Title == "Матн текшируви")
                        {
                            Require(!form.ContentPanel.HorizontalScroll.Visible && !form.ContentPanel.VerticalScroll.Visible,
                                "Short success message needs no scrollbars");
                            using (var screenshot = new Bitmap(form.Width, form.Height))
                            {
                                form.DrawToBitmap(screenshot, form.ClientRectangle);
                                screenshot.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SuccessMessage-" + dpi + ".png"));
                            }
                        }
                        Require(Math.Abs(form.Font.Size - ThemeManager.FontBase.SizeInPoints * dpi / 72f) < 0.1,
                            "Physical font size follows DPI");
                        if (form is ControlTestForm test)
                        {
                            Require(test.Grid.RowTemplate.Height == 40 * dpi / 96, "Grid row template scales");
                            Require(test.Grid.Rows[0].Height == 40 * dpi / 96, "Existing grid row scales");
                            Require(test.List.ItemHeight == 32 * dpi / 96, "Owner-drawn list row scales");
                        }
                        form.MinimumSize = Size.Empty;
                        form.Size = new Size(800, 560);
                        form.PerformLayout();
                        ResetScroll(form);
                        foreach (var button in Descendants(form).OfType<ModernButton>())
                            Require(button.Width >= ScreenGeometry.Scale(40, dpi),
                                "Button is not collapsed: " + form.GetType().Name + " / " + button.Text);
                        Require(form.ContentPanel.Parent is ScrollableControl, "Content scroll viewport");
                        Require(form.ActionBar.Parent is ScrollableControl, "Action scroll viewport");
                        Require(form.ContentPanel.Width >= form.ContentPanel.Parent.ClientSize.Width,
                            "Content is not crushed: " + form.GetType().Name + " " + form.ContentPanel.Size + " / " + form.ContentPanel.Parent.ClientSize);
                        Require(form.ActionBar.Height == ScreenGeometry.Scale(footerHeight, dpi), "Action bar remains sized");
                        using (var bitmap = new Bitmap(form.Width, form.Height))
                        {
                            form.DrawToBitmap(bitmap, form.ClientRectangle);
                            if (dpi == 96 || dpi == 192)
                                bitmap.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                    form.GetType().Name + "-" + dpi + ".png"));
                        }
                        var footer = (ScrollableControl)form.ActionBar.Parent;
                        if (footerHeight > 0 && form.ActionBar.Width > footer.ClientSize.Width)
                        {
                            footer.AutoScrollPosition = new Point(form.ActionBar.Width, 0);
                            Require(footer.AutoScrollPosition.X < 0, "Off-screen actions can be scrolled into view");
                        }
                        foreach (var largeSize in new[] { new Size(3840, 2120), new Size(7680, 4280) })
                        {
                            form.Size = largeSize;
                            form.PerformLayout();
                            Require(form.ActionBar.Width >= form.ActionBar.Parent.ClientSize.Width, "Footer fills large work area");
                            var resizedHost = (ScrollableControl)form.ContentPanel.Parent;
                            if (resizedHost.AutoScrollMinSize.IsEmpty)
                                Require(!resizedHost.HorizontalScroll.Visible && !resizedHost.VerticalScroll.Visible,
                                    "No stale scrollbars after enlarging the dialog");
                        }
                        scale.Invoke(form, new object[] { 96 });
                        Require(form.ActionBar.Height == footerHeight, "Round trip restores action height");
                        Console.WriteLine("PASS: " + form.GetType().Name + " at " + dpi + " DPI");
                    }
                }
                foreach (Rectangle area in new[] { new Rectangle(0, 0, 800, 560), new Rectangle(-1920, -1080, 1920, 1040), new Rectangle(3840, 0, 7680, 4280) })
                {
                    Rectangle fitted = ScreenGeometry.Fit(new Rectangle(-10000, 10000, 20000, 20000), area);
                    Require(area.Contains(fitted), "Bounds contained on monitor");
                }
                Console.WriteLine("PASS: small, negative-coordinate and 8K monitor bounds");
            }
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    [DllImport("user32.dll")] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr handle);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);

    private static void CheckInitialGridWindow()
    {
        foreach (int context in new[] { -1, -2, -4 })
        {
            var previous = SetThreadDpiAwarenessContext(new IntPtr(context));
            try
            {
                var factories = new Func<ModernForm>[]
                {
                    () => new TranslitExceptionsForm(
                        () => Enumerable.Range(1, 39).Select(i => new TranslitException("Example " + i, "Example " + i)).ToList(),
                        item => { }, item => { }, item => { }, () => { }),
                    () => new ViewErrorsForm(
                        Enumerable.Range(1, 76).Select(i => new ErrorEntry { Word = "Error " + i, ParagraphIndex = i }).ToList(),
                        new List<ErrorEntry> { new ErrorEntry { Word = "Grammar example" } },
                        (item, replacement) => { }, item => { }, item => AddWordResult.Invalid)
                };
                foreach (var factory in factories)
                using (var form = factory())
                {
                    form.Opacity = 0;
                    Exception failure = null;
                    form.Shown += (s, e) => form.BeginInvoke((Action)(() =>
                    {
                        try
                        {
                            CheckNativeGrid(form, form is ViewErrorsForm ? 76 : 39);
                            if (form is ViewErrorsForm)
                            {
                                var select = typeof(ViewErrorsForm).GetMethod("SelectPill", BindingFlags.Instance | BindingFlags.NonPublic);
                                select.Invoke(form, new object[] { 1 });
                                CheckNativeGrid(form, 1);
                                select.Invoke(form, new object[] { 0 });
                                CheckNativeGrid(form, 76);
                            }
                        }
                        catch (Exception ex) { failure = ex; }
                        finally { form.Close(); }
                    }));
                    form.ShowDialog();
                    if (failure != null) throw failure;
                    Console.WriteLine("PASS: initial modal " + form.GetType().Name + " in DPI context " + context);
                }
            }
            finally { SetThreadDpiAwarenessContext(previous); }
        }
    }

    private static void CheckNativeGrid(ModernForm form, int expectedRows)
    {
        var grid = Descendants(form).OfType<DataGridView>().Single(g => g.Visible);
        Require(GetParent(grid.Handle) == grid.Parent.Handle, "Initial native grid parent matches managed parent");
        Require(IsWindowVisible(grid.Handle), "Initial native grid is visible before repaint or resize");
        Require(grid.Height > 100 && grid.Width > 100, "Initial grid has usable bounds");
        Require(grid.Rows.Count == expectedRows, "Expected rows are present");
    }

    private static void CheckDeferredErrorSuggestions()
    {
        using (new DpiLayout.Context())
        {
            int calls = 0;
            var errors = Enumerable.Range(1, 76).Select(i => new ErrorEntry
            {
                Word = "Error " + i,
                SuggestionsProvider = word =>
                {
                    calls++;
                    return new List<Suggestion> { new Suggestion { Text = "Correction" } };
                }
            }).ToList();
            using (var form = new ViewErrorsForm(errors, new List<ErrorEntry>(),
                (entry, text) => { }, entry => { }, entry => AddWordResult.Invalid))
            {
                Require(calls == 0, "Constructing error dialog must not calculate suggestions for all rows");
                form.Opacity = 0;
                form.Show();
                Application.DoEvents();
                Require(calls == 1, "Opening error dialog calculates only the selected error");
                var grid = Descendants(form).OfType<DataGridView>().Single(g => g.Visible);
                grid.ClearSelection();
                grid.Rows[1].Selected = true;
                Require(calls == 2, "Selecting another error calculates its suggestions on demand");
                grid.ClearSelection();
                grid.Rows[0].Selected = true;
                Require(calls == 2, "Revisiting a row reuses cached suggestions");
                var scale = typeof(DpiForm).GetMethod("ApplyDpi", BindingFlags.Instance | BindingFlags.NonPublic);
                foreach (int dpi in new[] { 96, 120, 144, 192, 240, 288, 384 })
                {
                    scale.Invoke(form, new object[] { dpi });
                    form.MinimumSize = Size.Empty;
                    form.Size = new Size(ScreenGeometry.Scale(800, dpi), ScreenGeometry.Scale(600, dpi));
                    form.PerformLayout();
                    var footer = (ScrollableControl)form.ActionBar.Parent;
                    Require(!footer.HorizontalScroll.Visible && !footer.VerticalScroll.Visible,
                        "Error dialog footer fits an 800-logical-pixel window at " + dpi);
                    foreach (var button in Descendants(form.ActionBar).OfType<ModernButton>().Where(b => b.Visible))
                    {
                        var bounds = form.ActionBar.RectangleToClient(button.RectangleToScreen(button.ClientRectangle));
                        Require(form.ActionBar.ClientRectangle.Contains(bounds), "Footer button stays reachable: " + button.Text);
                    }
                }
                Require(calls == 2, "Resizing must not load suggestions for unselected rows");
            }
            Console.WriteLine("PASS: lazy error suggestions, cached selection, and compact footer at 100-400%");
        }
    }

    private static void CheckExceptionFonts()
    {
        using (new DpiLayout.Context())
        using (var form = new TranslitExceptionsForm(
            () => new List<TranslitException> { new TranslitException("Microsoft", "Microsoft", "Brand") },
            item => { }, item => { }, item => { }, () => { }))
        {
            form.Opacity = 0;
            form.Show();
            var grid = Descendants(form).OfType<DataGridView>().Single();
            Application.DoEvents();
            var scale = typeof(DpiForm).GetMethod("ApplyDpi", BindingFlags.Instance | BindingFlags.NonPublic);
            foreach (int dpi in new[] { 120, 192, 144, 384, 120, 96 })
            {
                scale.Invoke(form, new object[] { dpi });
                var font = grid.Rows[0].Cells[1].InheritedStyle.Font;
                Require(font.Unit == GraphicsUnit.Pixel && Math.Abs(font.Size - 11f * dpi / 72f) < 0.1f,
                    "Brand-name font maintains its 11pt logical size across DPI changes");
            }
            Console.WriteLine("PASS: exception brand fonts across repeated DPI transitions");
        }
    }

    private static void ResetScroll(Control root)
    {
        if (root is ScrollableControl viewport) viewport.AutoScrollPosition = Point.Empty;
        foreach (Control child in root.Controls) ResetScroll(child);
    }

    private static System.Collections.Generic.IEnumerable<Control> Descendants(Control root)
    {
        yield return root;
        foreach (Control child in root.Controls)
            foreach (var item in Descendants(child)) yield return item;
    }

    private sealed class ControlTestForm : ModernForm
    {
        internal readonly DataGridView Grid;
        internal readonly ListBox List;
        internal ControlTestForm()
        {
            Size = new Size(900, 650);
            Grid = new DataGridView { Location = new Point(20, 140), Size = new Size(400, 160) };
            Grid.RowTemplate.Height = 40;
            Grid.Columns.Add("Text", "Text");
            Grid.Rows.Add("Ўзбекча матн");
            List = new ListBox { Location = new Point(450, 140), Size = new Size(200, 160),
                DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = 32 };
            ContentPanel.Controls.AddRange(new Control[] { Grid, List,
                new ModernTextBox { Location = new Point(20, 20), Size = new Size(300, 44), Text = "Ўзбекча матн" },
                new ModernToggle { Location = new Point(20, 80), Size = new Size(300, 40), Text = "Enabled" }
            });
        }
    }
}
