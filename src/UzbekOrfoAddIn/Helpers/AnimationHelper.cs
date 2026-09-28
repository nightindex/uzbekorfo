using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Provides animation utilities for WinForms controls вЂ” fade, slide, scale.
    /// Used to give the modern UI a polished, responsive feel.
    /// Form properties must be updated from the thread that owns the control.
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
        /// Runs animation frames on the control's WinForms message thread.
        /// The returned task can complete on another thread, so callers should
        /// marshal any UI work they do after awaiting it.
        /// </summary>
        private static Task Animate(Control control, int durationMs, int steps,
            Action initialize, Action<int> render, Action finish)
        {
            return InvokeOnUiThreadAsync(control, () =>
            {
                if (control.IsDisposed) return Task.CompletedTask;

                var completion = new TaskCompletionSource<bool>(
                    TaskCreationOptions.RunContinuationsAsynchronously);
                Timer timer = null;
                try
                {
                    initialize?.Invoke();
                    int frame = 0;
                    timer = new Timer { Interval = Math.Max(1, durationMs / steps) };
                    timer.Tick += (sender, args) =>
                    {
                        try
                        {
                            if (control.IsDisposed)
                            {
                                timer.Stop();
                                timer.Dispose();
                                completion.TrySetResult(true);
                                return;
                            }

                            render(++frame);
                            if (frame < steps) return;

                            timer.Stop();
                            timer.Dispose();
                            finish?.Invoke();
                            completion.TrySetResult(true);
                        }
                        catch (Exception ex)
                        {
                            try { timer.Stop(); timer.Dispose(); } catch { }
                            completion.TrySetException(ex);
                        }
                    };
                    timer.Start();
                }
                catch (Exception ex)
                {
                    try { timer?.Dispose(); } catch { }
                    completion.TrySetException(ex);
                }
                return completion.Task;
            });
        }

        /// <summary>Fades a form in from transparent to opaque.</summary>
        public static Task FadeIn(Form form, int durationMs = 200)
        {
            const int steps = 20;
            return Animate(form, durationMs, steps,
                () => { form.Opacity = 0; form.Show(); },
                frame => form.Opacity = (double)frame / steps,
                () => form.Opacity = 1.0);
        }

        /// <summary>Fades a form out, then hides it.</summary>
        public static Task FadeOut(Form form, int durationMs = 150)
        {
            const int steps = 15;
            return Animate(form, durationMs, steps,
                null,
                frame => form.Opacity = (double)(steps - frame) / steps,
                () => form.Hide());
        }

        /// <summary>Slides a control horizontally into position.</summary>
        public static Task SlideIn(Control control, int fromXOffset, int durationMs = 300)
        {
            const int steps = 20;
            int targetX = 0;
            int startX = 0;
            return Animate(control, durationMs, steps,
                () =>
                {
                    targetX = control.Left;
                    startX = targetX + fromXOffset;
                    control.Left = startX;
                    control.Visible = true;
                },
                frame =>
                {
                    double progress = EaseOut((double)frame / steps);
                    control.Left = startX + (int)((targetX - startX) * progress);
                },
                () => control.Left = targetX);
        }

        /// <summary>Slides a control vertically into position.</summary>
        public static Task SlideInVertical(Control control, int fromYOffset, int durationMs = 250)
        {
            const int steps = 18;
            int targetY = 0;
            int startY = 0;
            return Animate(control, durationMs, steps,
                () =>
                {
                    targetY = control.Top;
                    startY = targetY + fromYOffset;
                    control.Top = startY;
                    control.Visible = true;
                },
                frame =>
                {
                    double progress = EaseOut((double)frame / steps);
                    control.Top = startY + (int)((targetY - startY) * progress);
                },
                () => control.Top = targetY);
        }

        /// <summary>Applies a subtle scale bounce effect.</summary>
        public static Task ScaleBounce(Control control, int durationMs = 300)
        {
            const int steps = 15;
            Size originalSize = Size.Empty;
            Point originalLocation = Point.Empty;
            return Animate(control, durationMs, steps,
                () =>
                {
                    originalSize = control.Size;
                    originalLocation = control.Location;
                },
                frame =>
                {
                    double t = (double)(frame - 1) / steps;
                    double scale = t < 0.4
                        ? 1.0 + 0.1 * (t / 0.4)
                        : 1.1 - 0.1 * ((t - 0.4) / 0.6);
                    int newW = (int)(originalSize.Width * scale);
                    int newH = (int)(originalSize.Height * scale);
                    int offsetX = (originalSize.Width - newW) / 2;
                    int offsetY = (originalSize.Height - newH) / 2;
                    control.Size = new Size(newW, newH);
                    control.Location = new Point(originalLocation.X + offsetX, originalLocation.Y + offsetY);
                },
                () =>
                {
                    control.Size = originalSize;
                    control.Location = originalLocation;
                });
        }

        /// <summary>Smoothly transitions a control's background color.</summary>
        public static Task ColorTransition(Control control, Color fromColor, Color toColor,
            int durationMs = 200)
        {
            const int steps = 15;
            return Animate(control, durationMs, steps,
                null,
                frame =>
                {
                    double t = (double)frame / steps;
                    int r = (int)(fromColor.R + (toColor.R - fromColor.R) * t);
                    int g = (int)(fromColor.G + (toColor.G - fromColor.G) * t);
                    int b = (int)(fromColor.B + (toColor.B - fromColor.B) * t);
                    control.BackColor = Color.FromArgb(
                        Math.Max(0, Math.Min(255, r)),
                        Math.Max(0, Math.Min(255, g)),
                        Math.Max(0, Math.Min(255, b)));
                },
                () => control.BackColor = toColor);
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
