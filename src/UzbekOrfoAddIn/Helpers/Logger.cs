using System;
using System.Collections.Concurrent;
using System.IO;
using System.Text;
using System.Threading;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Simple file-based logger for diagnostics and error tracking.
    /// Writes to %AppData%/UzbekOrfo/log.txt.
    /// Uses a background flush to avoid blocking the UI thread on every call.
    /// </summary>
    public static class Logger
    {
        private static readonly object _lock = new object();
        private static string _logPath;
        private const long MaxLogSizeBytes = 5 * 1024 * 1024; // 5 MB
        private const long KeepTailBytes = 2 * 1024 * 1024;   // keep last ~2 MB on rotation

        private static readonly ConcurrentQueue<string> _queue = new ConcurrentQueue<string>();
        private static Timer _flushTimer;
        private static int _flushing; // 0 = idle, 1 = flushing (interlocked guard)

        /// <summary>
        /// Path to the log file.
        /// </summary>
        public static string LogPath
        {
            get
            {
                if (_logPath == null)
                {
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var dir = Path.Combine(appData, "UzbekOrfo");
                    Directory.CreateDirectory(dir);
                    _logPath = Path.Combine(dir, "log.txt");
                }
                return _logPath;
            }
        }

        static Logger()
        {
            // Flush every 500 ms; keeps file I/O off the UI thread
            _flushTimer = new Timer(_ => Flush(), null, 500, 500);
        }

        /// <summary>
        /// Logs an informational message.
        /// </summary>
        public static void Info(string message) => Enqueue("INFO", message);

        /// <summary>
        /// Logs a warning message.
        /// </summary>
        public static void Warn(string message) => Enqueue("WARN", message);

        /// <summary>
        /// Logs an error message.
        /// </summary>
        public static void Error(string message) => Enqueue("ERROR", message);

        /// <summary>
        /// Logs an exception with its full stack trace.
        /// </summary>
        public static void Error(string message, Exception ex) =>
            Enqueue("ERROR", $"{message}\n  Exception: {ex.GetType().Name}: {ex.Message}\n  Stack: {ex.StackTrace}");

        private static void Enqueue(string level, string message)
        {
            try
            {
                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
                _queue.Enqueue(line);
            }
            catch
            {
                // Logging should never crash the application
            }
        }

        /// <summary>
        /// Flushes all queued log lines to disk in a single file operation.
        /// </summary>
        public static void Flush()
        {
            // Prevent concurrent flushes
            if (Interlocked.CompareExchange(ref _flushing, 1, 0) != 0) return;

            try
            {
                if (_queue.IsEmpty) return;

                var sb = new StringBuilder();
                while (_queue.TryDequeue(out var line))
                    sb.AppendLine(line);

                if (sb.Length == 0) return;

                lock (_lock)
                {
                    RotateIfNeeded();
                    File.AppendAllText(LogPath, sb.ToString());
                }
            }
            catch
            {
                // Logging should never crash the application
            }
            finally
            {
                Interlocked.Exchange(ref _flushing, 0);
            }
        }

        /// <summary>
        /// Truncates the log file if it exceeds MaxLogSizeBytes.
        /// Uses stream-based tail reading to avoid loading the whole file into memory.
        /// </summary>
        private static void RotateIfNeeded()
        {
            try
            {
                var fi = new FileInfo(LogPath);
                if (!fi.Exists || fi.Length < MaxLogSizeBytes) return;

                long seekPos = fi.Length - KeepTailBytes;
                if (seekPos < 0) seekPos = 0;

                byte[] tailBytes;
                using (var fs = new FileStream(LogPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    fs.Seek(seekPos, SeekOrigin.Begin);
                    // Advance to next newline so we don't cut mid-line
                    int b;
                    while ((b = fs.ReadByte()) != -1 && b != '\n') { }

                    int remaining = (int)(fs.Length - fs.Position);
                    tailBytes = new byte[remaining];
                    int read = 0;
                    while (read < remaining)
                    {
                        int n = fs.Read(tailBytes, read, remaining - read);
                        if (n == 0) break;
                        read += n;
                    }
                }

                var header = Encoding.UTF8.GetBytes(
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [INFO] === Log rotated (was {fi.Length / 1024}KB) ==={Environment.NewLine}");

                using (var fs = new FileStream(LogPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    fs.Write(header, 0, header.Length);
                    fs.Write(tailBytes, 0, tailBytes.Length);
                }
            }
            catch { }
        }

        /// <summary>
        /// Clears the log file.
        /// </summary>
        public static void Clear()
        {
            try
            {
                // Drain queue first
                while (_queue.TryDequeue(out _)) { }

                lock (_lock)
                {
                    File.WriteAllText(LogPath, string.Empty);
                }
            }
            catch { }
        }
    }
}
