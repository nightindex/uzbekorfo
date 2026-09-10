using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Windows.Forms;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    /// <summary>
    /// Modern progress dialog for long-running import operations (folder import).
    /// Runs the import on an STA thread (required for COM interop with Word/Excel)
    /// and shows file-by-file progress with cancellation support.
    /// </summary>
    public class ImportProgressForm : ModernForm
    {
        private Label _statusLabel;
        private Label _detailLabel;
        private ProgressBarCustom _progressBar;
        private ModernButton _cancelButton;

        private Thread _workerThread;
        private volatile bool _cancelRequested;
        private volatile bool _workCompleted;

        /// <summary>Set to true if the user cancelled.</summary>
        public bool WasCancelled { get; private set; }

        /// <summary>Result populated after successful completion.</summary>
        public object WorkerResult { get; private set; }

        /// <summary>Exception if the background work faulted.</summary>
        public Exception WorkerError { get; private set; }

        /// <summary>Delegate for the work to perform on the background STA thread.</summary>
        public delegate void ImportWorkHandler(ImportProgressForm progress);

        public ImportProgressForm()
        {
            Title = "Импорт жараёни";
            Size = new Size(520, 280);
            MinimumSize = new Size(440, 240);
            ShowMinimizeButton = false;
            AllowResize = false;

            BuildUI();
        }

        private void BuildUI()
        {
            ContentPanel.Padding = new Padding(ThemeManager.SpaceXL, ThemeManager.SpaceLG,
                                               ThemeManager.SpaceXL, ThemeManager.SpaceLG);

            _statusLabel = new Label
            {
                Text = "Импорт бошланмоқда...",
                AutoSize = true,
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = ThemeManager.TextPrimary,
                Margin = new Padding(0, 0, 0, 8)
            };

            _detailLabel = new Label
            {
                Text = "",
                AutoSize = true,
                Font = new Font("Segoe UI", 11f, FontStyle.Regular),
                ForeColor = ThemeManager.TextSecondary,
                Margin = new Padding(0, 0, 0, 16)
            };

            _progressBar = new ProgressBarCustom
            {
                Height = 14,
                Dock = DockStyle.Top,
                Margin = new Padding(0, 0, 0, 12),
                Value = 0,
                Maximum = 100
            };

            var stack = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(0)
            };

            // Set widths so they fill properly
            _statusLabel.MaximumSize = new Size(440, 0);
            _detailLabel.MaximumSize = new Size(440, 0);
            _progressBar.Width = 440;

            stack.Controls.Add(_statusLabel);
            stack.Controls.Add(_detailLabel);
            stack.Controls.Add(_progressBar);

            ContentPanel.Controls.Add(stack);

            // Cancel button in action bar
            _cancelButton = new ModernButton
            {
                Text = "Бекор қилиш",
                Style = ModernButton.ButtonStyle.Secondary,
                Font = ThemeManager.FontLGBold,
                Size = new Size(170, 46),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            _cancelButton.Location = new Point(ActionBar.Width - 170 - ThemeManager.SpaceXL, 6);
            _cancelButton.Click += (s, e) =>
            {
                _cancelRequested = true;
                WasCancelled = true;
                _cancelButton.Enabled = false;
                _cancelButton.Text = "Тўхтатилмоқда...";
                _statusLabel.Text = "Тўхтатилмоқда...";
            };
            ActionBar.Controls.Add(_cancelButton);
        }

        /// <summary>
        /// Shows the dialog and executes the given work on a dedicated STA thread.
        /// The STA thread is required because file import may use COM interop
        /// (Word/Excel) which only works on STA threads.
        /// Returns DialogResult.OK on success, Cancel on user cancel, Abort on error.
        /// </summary>
        public DialogResult RunWithWork(ImportWorkHandler work)
        {
            Shown += (s, e) =>
            {
                _workerThread = new Thread(() =>
                {
                    try
                    {
                        work(this);
                        _workCompleted = true;
                        BeginInvokeOnUI(() =>
                        {
                            DialogResult = DialogResult.OK;
                        });
                    }
                    catch (ThreadAbortException)
                    {
                        _workCompleted = true;
                        Thread.ResetAbort();
                        BeginInvokeOnUI(() =>
                        {
                            WasCancelled = true;
                            DialogResult = DialogResult.Cancel;
                        });
                    }
                    catch (Exception ex)
                    {
                        _workCompleted = true;
                        WorkerError = ex;
                        BeginInvokeOnUI(() =>
                        {
                            DialogResult = DialogResult.Abort;
                        });
                    }
                });

                _workerThread.SetApartmentState(ApartmentState.STA);
                _workerThread.IsBackground = true;
                _workerThread.Name = "ImportWorker";
                _workerThread.Start();
            };

            return ShowDialog();
        }

        /// <summary>
        /// Report progress from the worker thread. Thread-safe — marshals to UI thread.
        /// </summary>
        public void ReportProgress(int percent, string status, string detail)
        {
            if (_workCompleted || IsDisposed) return;

            BeginInvokeOnUI(() =>
            {
                if (IsDisposed) return;
                _progressBar.Value = Math.Min(percent, _progressBar.Maximum);
                _statusLabel.Text = status ?? "";
                _detailLabel.Text = detail ?? "";
            });
        }

        /// <summary>Check whether cancellation was requested (thread-safe).</summary>
        public bool CancellationPending => _cancelRequested;

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Prevent closing while worker is running (unless cancellation is in progress)
            if (_workerThread != null && _workerThread.IsAlive && !_cancelRequested)
            {
                _cancelRequested = true;
                WasCancelled = true;
                e.Cancel = true;
                return;
            }
            base.OnFormClosing(e);
        }

        /// <summary>Helper to marshal a delegate to the UI thread safely.</summary>
        private void BeginInvokeOnUI(Action action)
        {
            if (IsDisposed || !IsHandleCreated) return;
            try
            {
                BeginInvoke(action);
            }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }
    }

    /// <summary>
    /// A custom-painted progress bar that matches the modern dark theme.
    /// </summary>
    internal class ProgressBarCustom : Control
    {
        private int _value;
        private int _maximum = 100;

        public int Value
        {
            get => _value;
            set { _value = Math.Max(0, Math.Min(value, _maximum)); Invalidate(); }
        }

        public int Maximum
        {
            get => _maximum;
            set { _maximum = Math.Max(1, value); Invalidate(); }
        }

        public ProgressBarCustom()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw |
                     ControlStyles.SupportsTransparentBackColor, true);

            Height = 12;
            BackColor = Color.Transparent;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int r = Height / 2;
            var trackRect = new Rectangle(0, 0, Width, Height);

            // Track background
            using (var trackPath = RoundedRect(trackRect, r))
            using (var trackBrush = new SolidBrush(ThemeManager.Surface))
            {
                g.FillPath(trackBrush, trackPath);
            }

            // Fill
            if (_value > 0 && _maximum > 0)
            {
                float fraction = (float)_value / _maximum;
                int fillWidth = Math.Max(Height, (int)(Width * fraction));
                var fillRect = new Rectangle(0, 0, fillWidth, Height);

                using (var fillPath = RoundedRect(fillRect, r))
                using (var fillBrush = new LinearGradientBrush(
                    fillRect, ThemeManager.Primary, ThemeManager.PrimaryHover, 0f))
                {
                    g.FillPath(fillBrush, fillPath);
                }
            }
        }

        private static GraphicsPath RoundedRect(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
