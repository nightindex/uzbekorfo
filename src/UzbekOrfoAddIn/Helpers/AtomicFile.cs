using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace UzbekOrfoAddIn.Helpers
{
    internal static class AtomicFile
    {
        /// <summary>Serialize writers across Word processes and preserve the previous file on failure.</summary>
        public static void WriteAllLines(string path, string[] lines)
        {
            string fullPath = Path.GetFullPath(path);
            string mutexName;
            using (var hash = SHA256.Create())
                mutexName = "Local\\UzbekOrfo.Settings." + BitConverter.ToString(
                    hash.ComputeHash(Encoding.UTF8.GetBytes(fullPath.ToUpperInvariant()))).Replace("-", "");

            using (var mutex = new Mutex(false, mutexName))
            {
                bool acquired = false;
                string temporary = fullPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    try { acquired = mutex.WaitOne(TimeSpan.FromSeconds(5)); }
                    catch (AbandonedMutexException) { acquired = true; }
                    if (!acquired) throw new IOException("Another Word instance is still saving settings.");
                    Directory.CreateDirectory(Path.GetDirectoryName(fullPath));
                    using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    {
                        using (var writer = new StreamWriter(stream, new UTF8Encoding(false), 1024, true))
                            foreach (string line in lines) writer.WriteLine(line);
                        stream.Flush(true);
                    }
                    if (File.Exists(fullPath))
                        File.Replace(temporary, fullPath, fullPath + ".bak");
                    else
                        File.Move(temporary, fullPath);
                }
                finally
                {
                    try { if (File.Exists(temporary)) File.Delete(temporary); }
                    finally { if (acquired) mutex.ReleaseMutex(); }
                }
            }
        }
    }
}
