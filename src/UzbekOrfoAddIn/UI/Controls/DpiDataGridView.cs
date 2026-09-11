using System.Windows.Forms;
using System;
using System.Runtime.InteropServices;

namespace UzbekOrfoAddIn.UI.Controls
{
    /// <summary>
    /// A grid can create its HWND while rows are populated, before its form is shown.
    /// Match the form's DPI context so Windows can parent that HWND correctly when
    /// the caller is an Office callback running in a different DPI context.
    /// </summary>
    internal sealed class DpiDataGridView : DataGridView
    {
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr handle);

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            // An early-created grid can be left in WinForms' hidden parking window
            // after the form acquires its final DPI context. Recreate against the
            // actual parent; do not confuse managed Visible with native visibility.
            if (Visible && IsHandleCreated && Parent != null && Parent.IsHandleCreated &&
                GetParent(Handle) != Parent.Handle)
                RecreateHandle();
        }

        protected override void CreateHandle()
        {
            if (Parent != null)
            {
                using (new DpiLayout.Context(DpiLayout.WindowContext(Parent))) base.CreateHandle();
            }
            else base.CreateHandle();
        }
    }
}
