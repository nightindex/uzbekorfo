using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Provides animation utilities for WinForms controls вЂ” fade, slide, scale.
    /// Used to give the modern UI a polished, responsive feel.
    /// All animations stay on the UI thread using WinForms Timer to avoid cross-thread issues.
    /// </summary>
    public static class AnimationHelper
    {
        private static Task InvokeOnUiThreadAsync(Control control, Func<Task> asyncAction)
        {
            if (control == null || control.IsDisposed) return Task.CompletedTask;
            if (!control.InvokeRequired) return asyncAction();

            var tcs = new TaskCompletionSource<bool>();
            try
            {
                control.BeginInvoke((MethodInvoker)(async delegate
                {
                    try
                    {
                        if (control.IsDisposed)
                        {
                            tcs.TrySetResult(true);
                            return;
                        }

                        await asyncAction();
                        tcs.TrySetResult(true);
                    }
                    catch (Exception ex)
                    {
                        tcs.TrySetException(ex);
                    }
                }));
            }
            catch (ObjectDisposedException)
            {
                tcs.TrySetResult(true);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }

            return tcs.Task;
        }

        /// <summary>
        /// UI-thread-safe delay using System.Windows.Forms.Timer.
        /// Unlike Task.Delay, this guarantees continuation on the UI thread.
        /// </summary>
        private static Task UIDelay(int milliseconds)
        {
            var tcs = new TaskCompletionSource<bool>();
            var timer = new Timer { Interval = Math.Max(1, milliseconds) };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                tcs.TrySetResult(true);
            };
            timer.Start();
            return tcs.Task;
        }
        /// <summary>
        /// Fades a form in from fully transparent to fully opaque.
        /// </summary>
        public static Task FadeIn(Form form, int durationMs = 200)
        {
            return InvokeOnUiThreadAsync(form, () => FadeInCore(form, durationMs));
        }

        private static async Task FadeInCore(Form form, int durationMs)
        {
            if (form == null || form.IsDisposed) return;

            form.Opacity = 0;
            form.Show();
            int steps = 20;
            int stepDelay = Math.Max(1, durationMs / steps);

            for (int i = 1; i <= steps; i++)
            {
                if (form.IsDisposed) return;
                form.Opacity = (double)i / steps;
                await UIDelay(stepDelay);
            }

            form.Opacity = 1.0;
        }

        /// <summary>
        /// Fades a form out from fully opaque to fully transparent, then hides it.
        /// </summary>
        public static Task FadeOut(Form form, int durationMs = 150)
        {
            return InvokeOnUiThreadAsync(form, () => FadeOutCore(form, durationMs));
        }

        private static async Task FadeOutCore(Form form, int durationMs)
        {
            if (form == null || form.IsDisposed) return;

            int steps = 15;
            int stepDelay = Math.Max(1, durationMs / steps);

            for (int i = steps - 1; i >= 0; i--)
            {
                if (form.IsDisposed) return;
                form.Opacity = (double)i / steps;
                await UIDelay(stepDelay);
            }

            form.Opacity = 0;
            form.Hide();
        }

        /// <summary>
        /// Slides a control into position from a given horizontal offset.
        /// </summary>
        public static Task SlideIn(Control control, int fromXOffset, int durationMs = 300)
        {
            return InvokeOnUiThreadAsync(control, () => SlideInCore(control, fromXOffset, durationMs));
        }

        private static async Task SlideInCore(Control control, int fromXOffset, int durationMs)
        {
            if (control == null || control.IsDisposed) return;

            int targetX = control.Left;
            int startX = targetX + fromXOffset;
            int steps = 20;
            int stepDelay = Math.Max(1, durationMs / steps);

            control.Left = startX;
            control.Visible = true;

            for (int i = 1; i <= steps; i++)
            {
                if (control.IsDisposed) return;
                double progress = EaseOut((double)i / steps);
                control.Left = startX + (int)((targetX - startX) * progress);
                await UIDelay(stepDelay);
            }

            control.Left = targetX;
        }

        /// <summary>
        /// Slides a control vertically into position from a given offset.
        /// </summary>
        public static Task SlideInVertical(Control control, int fromYOffset, int durationMs = 250)
        {
            return InvokeOnUiThreadAsync(control, () => SlideInVerticalCore(control, fromYOffset, durationMs));
        }

        private static async Task SlideInVerticalCore(Control control, int fromYOffset, int durationMs)
        {
            if (control == null || control.IsDisposed) return;

            int targetY = control.Top;
            int startY = targetY + fromYOffset;
            int steps = 18;
            int stepDelay = Math.Max(1, durationMs / steps);

            control.Top = startY;
            control.Visible = true;

            for (int i = 1; i <= steps; i++)
            {
                if (control.IsDisposed) return;
                double progress = EaseOut((double)i / steps);
                control.Top = startY + (int)((targetY - startY) * progress);
                await UIDelay(stepDelay);
            }

            control.Top = targetY;
        }

        /// <summary>
        /// Applies a subtle scale bounce effect (scale up briefly then back to normal).
        /// Simulated via size change since WinForms doesn't natively support transforms.
        /// </summary>
        public static Task ScaleBounce(Control control, int durationMs = 300)
        {
            return InvokeOnUiThreadAsync(control, () => ScaleBounceCore(control, durationMs));
        }

        private static async Task ScaleBounceCore(Control control, int durationMs)
        {
            if (control == null || control.IsDisposed) return;

            var originalSize = control.Size;
            var originalLocation = control.Location;
            int steps = 15;
            int stepDelay = Math.Max(1, durationMs / steps);

            for (int i = 0; i < steps; i++)
            {
                if (control.IsDisposed) return;
                double t = (double)i / steps;
                double scale;

                if (t < 0.4)
                    scale = 1.0 + 0.1 * (t / 0.4); // grow to 1.1
                else
                    scale = 1.1 - 0.1 * ((t - 0.4) / 0.6); // shrink back to 1.0

                int newW = (int)(originalSize.Width * scale);
                int newH = (int)(originalSize.Height * scale);
                int offsetX = (originalSize.Width - newW) / 2;
                int offsetY = (originalSize.Height - newH) / 2;

                control.Size = new Size(newW, newH);
                control.Location = new Point(originalLocation.X + offsetX, originalLocation.Y + offsetY);

                await UIDelay(stepDelay);
            }

            control.Size = originalSize;
            control.Location = originalLocation;
        }

        /// <summary>
        /// Smoothly transitions the background color of a control.
        /// </summary>
        public static Task ColorTransition(Control control, Color fromColor, Color toColor, int durationMs = 200)
        {
            return InvokeOnUiThreadAsync(control, () => ColorTransitionCore(control, fromColor, toColor, durationMs));
        }

        private static async Task ColorTransitionCore(Control control, Color fromColor, Color toColor, int durationMs)
        {
            if (control == null || control.IsDisposed) return;

            int steps = 15;
            int stepDelay = Math.Max(1, durationMs / steps);

            for (int i = 1; i <= steps; i++)
            {
                if (control.IsDisposed) return;
                double t = (double)i / steps;
                int r = (int)(fromColor.R + (toColor.R - fromColor.R) * t);
                int g = (int)(fromColor.G + (toColor.G - fromColor.G) * t);
                int b = (int)(fromColor.B + (toColor.B - fromColor.B) * t);
                control.BackColor = Color.FromArgb(
                    Math.Max(0, Math.Min(255, r)),
                    Math.Max(0, Math.Min(255, g)),
                    Math.Max(0, Math.Min(255, b)));
                await UIDelay(stepDelay);
            }

            control.BackColor = toColor;
        }

        // --- Easing Functions ---

        /// <summary>
        /// Cubic ease-out: fast start, slow finish.
        /// </summary>
        private static double EaseOut(double t)
        {
            return 1.0 - Math.Pow(1.0 - t, 3);
        }

        /// <summary>
        /// Cubic ease-in: slow start, fast finish.
        /// </summary>
        public static double EaseIn(double t)
        {
            return t * t * t;
        }

        /// <summary>
        /// Cubic ease-in-out: slow start - fast middle - slow end.
        /// </summary>
        public static double EaseInOut(double t)
        {
            return t < 0.5
                ? 4 * t * t * t
                : 1 - Math.Pow(-2 * t + 2, 3) / 2;
        }
    }
}
