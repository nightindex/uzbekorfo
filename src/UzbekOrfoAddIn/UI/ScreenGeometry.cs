using System;
using System.Drawing;

namespace UzbekOrfoAddIn.UI
{
    internal static class ScreenGeometry
    {
        internal static int Scale(int value, int dpi) =>
            (int)Math.Round(value * Math.Max(96, dpi) / 96.0);

        internal static Rectangle Fit(Rectangle bounds, Rectangle workArea)
        {
            int width = Math.Min(bounds.Width, workArea.Width);
            int height = Math.Min(bounds.Height, workArea.Height);
            return new Rectangle(
                Math.Max(workArea.Left, Math.Min(bounds.X, workArea.Right - width)),
                Math.Max(workArea.Top, Math.Min(bounds.Y, workArea.Bottom - height)),
                width, height);
        }
    }
}
