using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UzbekOrfoAddIn.UI
{
    /// <summary>
    /// Code-built forms use 96-DPI design coordinates. Scale explicitly because
    /// an add-in cannot opt Word into WinForms application.config DPI settings.
    /// </summary>
    public class DpiForm : Form
    {
        private readonly Dictionary<string, Font> fonts = new Dictionary<string, Font>();
        private bool loaded;
        private bool scaling;
        private Size designMinimum;
        public int LayoutDpi { get; private set; } = 96;
        protected int Px(int value) => ScreenGeometry.Scale(value, LayoutDpi);
        protected Font UiFont(Font value) => GetDpiFont(value);
        internal Font GetDpiFont(Font value)
        {
            // Pixel fonts are already scaled by this form.
            if (value.Unit == GraphicsUnit.Pixel) return value;
            return PixelFont(value, value.SizeInPoints * LayoutDpi / 72f);
        }

        private Font PixelFont(Font value, float size)
        {
            // Reuse equivalent fonts when repeatedly moving between monitors.
            size = (float)Math.Round(size, 4);
            string key = value.Name + ":" + size + ":" + value.Style;
            if (!fonts.TryGetValue(key, out Font result))
            {
                result = new Font(value.FontFamily, size, value.Style, GraphicsUnit.Pixel);
                fonts.Add(key, result);
            }
            return result;
        }

        public DpiForm()
        {
            AutoScaleMode = AutoScaleMode.None;
        }

        protected override void CreateHandle()
        {
            using (new DpiLayout.Context()) base.CreateHandle();
        }

        protected override void OnLoad(EventArgs e)
        {
            using (new DpiLayout.Context())
            {
                if (!loaded)
                {
                    designMinimum = MinimumSize;
                    // Pick the invoking window's monitor before the first scale pass.
                    if (StartPosition != FormStartPosition.Manual)
                    {
                        var area = Owner != null ? Screen.FromControl(Owner).WorkingArea : DpiLayout.ActiveScreen().WorkingArea;
                        StartPosition = FormStartPosition.Manual;
                        Location = new Point(area.Left + (area.Width - Width) / 2, area.Top + (area.Height - Height) / 2);
                    }
                    loaded = true;
                    ApplyDpi(DpiLayout.WindowDpi(this));
                    FitToScreen();
                }
                base.OnLoad(e);
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            foreach (var viewport in Descendants(this).OfType<ScrollableControl>())
                viewport.AutoScrollPosition = Point.Empty;
        }

        protected virtual void OnLayoutDpiChanged() { }

        private static IEnumerable<Control> Descendants(Control root)
        {
            yield return root;
            foreach (Control child in root.Controls)
                foreach (Control item in Descendants(child)) yield return item;
        }

        private void ApplyDpi(int dpi)
        {
            if (scaling) return;
            scaling = true;
            var controls = Descendants(this).ToArray();
            foreach (var viewport in controls.OfType<ScrollableControl>())
                viewport.AutoScrollPosition = Point.Empty;
            var previousFonts = controls.ToDictionary(c => c, c => c.Font);
            var gridFonts = controls.OfType<DataGridView>().Select(grid => new
            {
                Grid = grid,
                Cell = grid.DefaultCellStyle.Font,
                Header = grid.ColumnHeadersDefaultCellStyle.Font,
                Alternate = grid.AlternatingRowsDefaultCellStyle.Font,
                RowHeader = grid.RowHeadersDefaultCellStyle.Font
            }).ToArray();
            var geometry = controls.Select(c => new
            {
                Control = c, c.Bounds, c.Padding, c.Margin, c.MinimumSize, c.MaximumSize,
                Rows = (c as TableLayoutPanel)?.RowStyles.Cast<RowStyle>().Select(s => s.Height).ToArray(),
                Columns = (c as TableLayoutPanel)?.ColumnStyles.Cast<ColumnStyle>().Select(s => s.Width).ToArray()
            }).ToArray();
            int oldDpi = LayoutDpi;
            float ratio = (float)dpi / oldDpi;
            foreach (var control in controls) control.SuspendLayout();
            try
            {
                LayoutDpi = dpi;
                MinimumSize = Size.Empty;
                // Control.Scale uses WinForms' deferred autoscaling/anchor state, which
                // can shrink table children to one pixel on subsequent DPI changes.
                // Snapshot all geometry first, then apply one explicit scale pass.
                foreach (var item in geometry)
                {
                    var control = item.Control;
                    control.MinimumSize = Size.Empty;
                    control.MaximumSize = Size.Empty;
                    control.Padding = ScalePadding(item.Padding, ratio);
                    control.Margin = ScalePadding(item.Margin, ratio);
                    var bounds = item.Bounds;
                    if (control == this) Size = ScaleSize(bounds.Size, ratio);
                    else control.Bounds = new Rectangle(
                        (int)Math.Round(bounds.X * ratio), (int)Math.Round(bounds.Y * ratio),
                        (int)Math.Round(bounds.Width * ratio), (int)Math.Round(bounds.Height * ratio));
                    if (control != this) control.MinimumSize = ScaleSize(item.MinimumSize, ratio);
                    control.MaximumSize = ScaleSize(item.MaximumSize, ratio);
                    if (control is TableLayoutPanel table)
                    {
                        for (int i = 0; i < item.Rows.Length; i++)
                            if (table.RowStyles[i].SizeType == SizeType.Absolute) table.RowStyles[i].Height = item.Rows[i] * ratio;
                        for (int i = 0; i < item.Columns.Length; i++)
                            if (table.ColumnStyles[i].SizeType == SizeType.Absolute) table.ColumnStyles[i].Width = item.Columns[i] * ratio;
                    }
                }
                foreach (var control in controls)
                {
                    Font original = previousFonts[control];
                    if (original.Unit == GraphicsUnit.Pixel)
                    {
                        control.Font = PixelFont(original, original.Size * ratio);
                    }
                    else control.Font = GetDpiFont(original);

                    if (control is ListBox list && list.DrawMode == DrawMode.OwnerDrawFixed)
                        list.ItemHeight = Math.Max(1, (int)Math.Round(list.ItemHeight * ratio));
                    if (control is DataGridView grid)
                    {
                        foreach (DataGridViewColumn column in grid.Columns)
                        {
                            column.MinimumWidth = Math.Max(2, (int)Math.Round(column.MinimumWidth * ratio));
                            if (column.InheritedAutoSizeMode == DataGridViewAutoSizeColumnMode.None)
                                column.Width = Math.Max(column.MinimumWidth, (int)Math.Round(column.Width * ratio));
                        }
                        grid.RowTemplate.Height = Math.Max(2, (int)Math.Round(grid.RowTemplate.Height * ratio));
                        grid.ColumnHeadersHeight = Math.Max(4, (int)Math.Round(grid.ColumnHeadersHeight * ratio));
                        foreach (DataGridViewRow row in grid.Rows)
                            row.Height = Math.Max(2, (int)Math.Round(row.Height * ratio));
                    }
                }
                OnLayoutDpiChanged();
            }
            finally
            {
                foreach (var control in controls.Reverse()) control.ResumeLayout(true);
                scaling = false;
            }
            // Native grid DPI/layout processing can replace its cell styles while
            // layouts resume. Apply fonts to the current styles AFTER that processing,
            // using the pre-change fonts rather than already adjusted native values.
            foreach (var item in gridFonts)
            {
                item.Grid.DefaultCellStyle = ScaleStyle(item.Grid.DefaultCellStyle, item.Cell, ratio);
                item.Grid.ColumnHeadersDefaultCellStyle = ScaleStyle(item.Grid.ColumnHeadersDefaultCellStyle, item.Header, ratio);
                item.Grid.AlternatingRowsDefaultCellStyle = ScaleStyle(item.Grid.AlternatingRowsDefaultCellStyle, item.Alternate, ratio);
                item.Grid.RowHeadersDefaultCellStyle = ScaleStyle(item.Grid.RowHeadersDefaultCellStyle, item.RowHeader, ratio);
            }
            Invalidate(true);
            foreach (var viewport in controls.OfType<ScrollableControl>())
                viewport.AutoScrollPosition = Point.Empty;
        }

        private DataGridViewCellStyle ScaleStyle(DataGridViewCellStyle style, Font original, float ratio)
        {
            // Replace the full style: mutating Font on the existing style can be
            // superseded by DataGridView's native font/style change notifications.
            // Cloning also preserves colors, alignment, formatting, and padding.
            style = new DataGridViewCellStyle(style);
            style.Font = original == null ? null : original.Unit != GraphicsUnit.Pixel
                ? GetDpiFont(original)
                : PixelFont(original, original.Size * ratio);
            return style;
        }

        private static Size ScaleSize(Size size, float ratio) =>
            new Size((int)Math.Round(size.Width * ratio), (int)Math.Round(size.Height * ratio));

        private static Padding ScalePadding(Padding padding, float ratio) =>
            new Padding((int)Math.Round(padding.Left * ratio), (int)Math.Round(padding.Top * ratio),
                (int)Math.Round(padding.Right * ratio), (int)Math.Round(padding.Bottom * ratio));

        protected void FitToScreen()
        {
            using (new DpiLayout.Context())
            {
                var area = Screen.FromControl(this).WorkingArea;
                MinimumSize = new Size(Math.Min(Px(designMinimum.Width), area.Width), Math.Min(Px(designMinimum.Height), area.Height));
                Bounds = ScreenGeometry.Fit(Bounds, area);
            }
        }

        protected override void WndProc(ref Message m)
        {
            const int WmDpiChanged = 0x02E0;
            if (m.Msg == WmDpiChanged && loaded && !scaling)
            {
                int dpi = (int)(m.WParam.ToInt64() & 0xffff);
                var suggested = (NativeRect)Marshal.PtrToStructure(m.LParam, typeof(NativeRect));
                ApplyDpi(dpi);
                Bounds = Rectangle.FromLTRB(suggested.Left, suggested.Top, suggested.Right, suggested.Bottom);
                FitToScreen();
                m.Result = IntPtr.Zero;
                return; // Explicit scaling: do not also let WinForms scale this message.
            }
            base.WndProc(ref m);
            // Also handle same-DPI monitors with different work areas and display removal.
            if (loaded && !scaling && (m.Msg == 0x0232 || m.Msg == 0x007E))
                FitToScreen();
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect { public int Left, Top, Right, Bottom; }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                foreach (Font font in fonts.Values) font.Dispose();
                fonts.Clear();
            }
        }
    }
}
