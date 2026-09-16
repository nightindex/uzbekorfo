using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using UzbekOrfoAddIn.Services;
using Word = Microsoft.Office.Interop.Word;

// Opt-in visible Word test; only captures a crop of its own disposable document.
internal static class GhostVisualChecks
{
    private sealed class Owner : IWin32Window
    {
        public IntPtr Handle { get; }
        internal Owner(int handle) { Handle = new IntPtr(handle); }
    }

    internal static void Run(Word.Application app, Word.Document document)
    {
        var assembly = typeof(WordCompletionContext).Assembly;
        var ghostType = assembly.GetType("UzbekOrfoAddIn.Forms.GhostSuggestionWindow", true);
        var getAnchor = typeof(WordCompletionContext).GetMethod("GetCaretAnchor", BindingFlags.Instance | BindingFlags.NonPublic);
        var present = ghostType.GetMethod("Present", BindingFlags.Instance | BindingFlags.NonPublic);
        using (var ghost = (Form)Activator.CreateInstance(ghostType, true))
        {
            app.Visible = true;
            app.Activate();
            document.Activate();
            foreach (string font in new[] { "Arial", "Calibri", "Times New Roman" })
            foreach (int zoom in new[] { 100, 150 })
            foreach (string prefix in new[] { "Сал", "Sal" })
            foreach (bool useFallback in new[] { false, true })
            {
                document.TrackRevisions = false;
                document.Content.Text = prefix;
                document.Content.Font.Name = font;
                document.Content.Font.Size = 14;
                app.ActiveWindow.View.Zoom.Percentage = zoom;
                document.Range(prefix.Length, prefix.Length).Select();
                app.ActiveWindow.ScrollIntoView(app.Selection.Range);
                Pump();
                string word = prefix == "Сал" ? "Салом" : "Salom";
                using (var context = WordCompletionContext.Capture(app, false))
                {
                    if (context == null) throw new InvalidOperationException("Cannot capture visual test prefix.");
                    // Suppress any installed add-in's automatic suggestions in this
                    // temporary document so they cannot overlap the renderer under test.
                    document.TrackRevisions = true;
                    Pump();
                    object anchor = getAnchor.Invoke(context, new object[] { app });
                    var flags = BindingFlags.Instance | BindingFlags.NonPublic;
                    int px, py, pw, ph;
                    app.ActiveWindow.GetPoint(out px, out py, out pw, out ph, document.Range(2, 3));
                    object[] nativeArgs = { new IntPtr(app.ActiveWindow.Hwnd), Rectangle.Empty };
                    bool native = (bool)assembly.GetType("UzbekOrfoAddIn.UI.OverlayPositioner").GetMethod(
                        "TryGetNativeCaret", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, nativeArgs);
                    Console.WriteLine("ANCHOR caret=" + anchor.GetType().GetProperty("CaretBounds", flags).GetValue(anchor) +
                        "; lastChar=" + new Rectangle(px, py, pw, ph) + "; native=" + native + " " + nativeArgs[1]);
                    if (useFallback)
                    {
                        var positioner = assembly.GetType("UzbekOrfoAddIn.UI.OverlayPositioner");
                        int cx, cy, cw, ch;
                        app.ActiveWindow.GetPoint(out cx, out cy, out cw, out ch, document.Range(3, 3));
                        anchor = positioner.GetMethod("CreateAnchor", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null,
                            new object[] { new Rectangle(cx, py, 1, ph), new IntPtr(app.ActiveWindow.Hwnd),
                                font, 14f * zoom / 100f, FontStyle.Regular });
                        anchor = positioner.GetMethod("WithTextHeight", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null,
                            new[] { anchor });
                    }
                    if (!(bool)present.Invoke(ghost, new object[] {
                        WordCompletionEngine.GetGhostTail(prefix, word), anchor, new Owner(app.ActiveWindow.Hwnd) }))
                        throw new InvalidOperationException("Ghost could not render in the Word document.");
                    Pump();
                    var bounds = (Rectangle)ghostType.GetProperty("PresentationBounds", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(ghost);
                    var crop = new Rectangle(bounds.Left, bounds.Top - 40, bounds.Width + 16, bounds.Height + 80);
                    string name = font.Replace(" ", "-") + "-" + zoom + "-" + (prefix == "Сал" ? "Cyrl" : "Latin") +
                        (useFallback ? "-fallback" : "-caret");
                    int wordX, wordY, wordWidth, wordHeight;
                    app.ActiveWindow.GetPoint(out wordX, out wordY, out wordWidth, out wordHeight, document.Range(0, 3));
                    using (var wordPreview = Capture(new Rectangle(wordX - 10, bounds.Top - 12,
                        bounds.Right - wordX + 20, bounds.Height + 24)))
                        wordPreview.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Word-" + name + ".png"), ImageFormat.Png);
                    using (var preview = Capture(crop))
                    {
                        preview.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ghost-" + name + ".png"), ImageFormat.Png);
                        ghost.Hide();
                        document.TrackRevisions = false;
                        if (!context.TryInsert(app, word, false)) throw new InvalidOperationException("Visual test acceptance failed.");
                        document.TrackRevisions = true;
                        // Move the caret away from the measured tail without changing the text.
                        document.Range(0, 0).Select();
                        Pump();
                        using (var accepted = Capture(crop))
                        {
                            accepted.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Accepted-" + name + ".png"), ImageFormat.Png);
                            Rectangle previewInk = InkBounds(preview), acceptedInk = InkBounds(accepted);
                            Console.WriteLine("GHOST " + name + ": preview=" + previewInk + "; accepted=" + acceptedInk);
                            if (previewInk.IsEmpty || acceptedInk.IsEmpty ||
                                Math.Abs(previewInk.Bottom - acceptedInk.Bottom) > 3 ||
                                Math.Abs(previewInk.Top - acceptedInk.Top) > 3)
                                throw new InvalidOperationException("Ghost text does not share Word's text baseline/size: " + name);
                        }
                    }
                }
            }
        }
        Console.WriteLine("PASS: ghost baseline and font size match accepted Word text in Latin/Cyrillic, three fonts, 100%/150% zoom.");
    }

    private static Bitmap Capture(Rectangle crop)
    {
        var bitmap = new Bitmap(crop.Width, crop.Height);
        using (var graphics = Graphics.FromImage(bitmap))
            graphics.CopyFromScreen(crop.Location, Point.Empty, crop.Size);
        return bitmap;
    }

    private static Rectangle InkBounds(Bitmap bitmap)
    {
        Color background = bitmap.GetPixel(bitmap.Width - 1, 0);
        int left = bitmap.Width, top = bitmap.Height, right = -1, bottom = -1;
        for (int y = 0; y < bitmap.Height; y++)
        for (int x = 2; x < bitmap.Width; x++)
        {
            Color pixel = bitmap.GetPixel(x, y);
            if (Math.Abs(pixel.R - background.R) + Math.Abs(pixel.G - background.G) + Math.Abs(pixel.B - background.B) < 120) continue;
            left = Math.Min(left, x); top = Math.Min(top, y);
            right = Math.Max(right, x); bottom = Math.Max(bottom, y);
        }
        return right < left ? Rectangle.Empty : Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
    }

    private static void Pump()
    {
        var watch = Stopwatch.StartNew();
        while (watch.ElapsedMilliseconds < 160)
        {
            Application.DoEvents();
            System.Threading.Thread.Sleep(10);
        }
    }
}
