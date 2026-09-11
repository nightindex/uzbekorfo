using System;
using System.IO;
using System.Reflection;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Copies bundled seed data files from the add-in folder into the runtime
    /// AppData directory when user-local files are missing, while refreshing
    /// explicitly built-in artifacts when a bundled release changes them.
    /// Supports both file-based (F5 debug) and embedded resource (ClickOnce) deployment.
    /// </summary>
    public static class DataSeedService
    {
        public static void SeedDictionary(string mainDictionaryPath)
        {
            if (string.IsNullOrWhiteSpace(mainDictionaryPath)) return;
            try
            {
                // The built-in dictionary has no user edits; update it when a
                // release ships a changed generated DIC. User words remain in
                // user_custom.dic and are never overwritten here.
                if (!EnsureBundledDataFileCurrent("uzbek_main.dic", mainDictionaryPath, out _))
                {
                    EnsureEmbeddedResourceCurrent("UzbekOrfoAddIn.Data.uzbek_main.dic", mainDictionaryPath);
                }
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error("Failed to seed dictionary", ex);
            }
        }

        /// <summary>
        /// Makes the compact generated built-in metadata index available when
        /// a definition lookup needs it. It is deliberately not read at startup.
        /// </summary>
        public static void SeedDictionaryMetadata(string dictionaryMetadataPath)
        {
            if (string.IsNullOrWhiteSpace(dictionaryMetadataPath)) return;
            try
            {
                if (!EnsureBundledDataFileCurrent(
                    "uzbek_dictionary_metadata.json", dictionaryMetadataPath, out _))
                {
                    EnsureEmbeddedResourceCurrent(
                        "UzbekOrfoAddIn.Data.uzbek_dictionary_metadata.json", dictionaryMetadataPath);
                }
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error("Failed to seed dictionary metadata", ex);
            }
        }

        public static void SeedTranslitExceptions(string exceptionsPath, TransliterationService transliterator)
        {
            if (string.IsNullOrWhiteSpace(exceptionsPath)) return;
            try
            {
                bool copied = EnsureFileWithMinimumSize(
                    exceptionsPath,
                    minimumBytes: 10,
                    bundledFileName: "translit_exceptions.json",
                    embeddedResourceName: "UzbekOrfoAddIn.Data.translit_exceptions.json");

                if (copied && transliterator != null)
                    transliterator.LoadExceptions();
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error("Failed to seed transliteration exceptions", ex);
            }
        }

        public static void SeedExplanations(string explanationsPath)
        {
            if (string.IsNullOrWhiteSpace(explanationsPath)) return;
            try
            {
                EnsureFileWithMinimumSize(
                    explanationsPath,
                    minimumBytes: 10,
                    bundledFileName: "explanations.json",
                    embeddedResourceName: "UzbekOrfoAddIn.Data.explanations.json");
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error("Failed to seed explanations", ex);
            }
        }

        public static void SeedGrammarFiles(string suffixesPath, string grammarRulesPath, string properNounsPath)
        {
            try
            {
                SeedIfMissing(suffixesPath, "uzbek_suffixes.json", "UzbekOrfoAddIn.Data.uzbek_suffixes.json");
                SeedIfMissing(grammarRulesPath, "grammar_rules.json", "UzbekOrfoAddIn.Data.grammar_rules.json");
                SeedIfMissing(properNounsPath, "proper_nouns.json", "UzbekOrfoAddIn.Data.proper_nouns.json");
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error("Failed to seed grammar files", ex);
            }
        }

        private static void SeedIfMissing(string targetPath, string bundledFileName, string embeddedResourceName)
        {
            if (string.IsNullOrWhiteSpace(targetPath)) return;
            if (File.Exists(targetPath)) return;
            
            if (!CopyBundledDataFile(bundledFileName, targetPath, overwrite: false))
            {
                ExtractEmbeddedResource(embeddedResourceName, targetPath);
            }
        }

        /// <summary>
        /// Finds the deployed data file and replaces the local built-in copy
        /// only when its bytes differ. This lets data updates reach existing
        /// installations without ever replacing user-owned files.
        /// </summary>
        private static bool EnsureBundledDataFileCurrent(string fileName, string targetPath, out bool updated)
        {
            updated = false;
            foreach (var sourcePath in GetSearchPaths(fileName))
            {
                if (!File.Exists(sourcePath)) continue;

                if (FilesEqual(sourcePath, targetPath))
                    return true;

                string targetDir = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrWhiteSpace(targetDir))
                    Directory.CreateDirectory(targetDir);

                File.Copy(sourcePath, targetPath, overwrite: true);
                updated = true;
                Helpers.Logger.Info($"Updated bundled file from: {sourcePath}");
                return true;
            }

            return false;
        }

        private static bool EnsureFileWithMinimumSize(string targetPath, long minimumBytes, string bundledFileName, string embeddedResourceName)
        {
            if (File.Exists(targetPath) && new FileInfo(targetPath).Length > minimumBytes)
                return false;

            if (CopyBundledDataFile(bundledFileName, targetPath, overwrite: true))
                return true;
                
            return ExtractEmbeddedResource(embeddedResourceName, targetPath);
        }

        private static bool CopyBundledDataFile(string fileName, string targetPath, bool overwrite)
        {
            // Try multiple locations for maximum compatibility
            string[] searchPaths = GetSearchPaths(fileName);
            
            foreach (var sourcePath in searchPaths)
            {
                if (File.Exists(sourcePath))
                {
                    string targetDir = Path.GetDirectoryName(targetPath);
                    if (!string.IsNullOrWhiteSpace(targetDir))
                        Directory.CreateDirectory(targetDir);

                    File.Copy(sourcePath, targetPath, overwrite);
                    Helpers.Logger.Info($"Seeded file from: {sourcePath}");
                    return true;
                }
            }

            Helpers.Logger.Warn($"Bundled seed file not found in any location: {fileName}");
            return false;
        }

        private static string[] GetSearchPaths(string fileName)
        {
            var paths = new System.Collections.Generic.List<string>();
            
            // 1. Assembly location (works for F5 debug and most deployments)
            string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
            paths.Add(Path.Combine(assemblyDir, "Data", fileName));
            paths.Add(Path.Combine(assemblyDir, fileName));
            
            // 2. CodeBase (sometimes different from Location in ClickOnce)
            try
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                if (!string.IsNullOrEmpty(codeBase))
                {
                    string codeBaseDir = Path.GetDirectoryName(new Uri(codeBase).LocalPath) ?? string.Empty;
                    if (!string.Equals(codeBaseDir, assemblyDir, StringComparison.OrdinalIgnoreCase))
                    {
                        paths.Add(Path.Combine(codeBaseDir, "Data", fileName));
                        paths.Add(Path.Combine(codeBaseDir, fileName));
                    }
                }
            }
            catch { }
            
            // 3. Current directory fallback
            paths.Add(Path.Combine(Environment.CurrentDirectory, "Data", fileName));
            
            return paths.ToArray();
        }

        /// <summary>
        /// Extracts an embedded resource to the target path.
        /// This is the most reliable method for ClickOnce deployments.
        /// </summary>
        private static bool ExtractEmbeddedResource(string resourceName, string targetPath)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        // Try case-insensitive search
                        var segments = resourceName.Split('.');
                        var lastSegment = segments[segments.Length - 1];
                        foreach (var name in assembly.GetManifestResourceNames())
                        {
                            if (name.EndsWith(lastSegment, StringComparison.OrdinalIgnoreCase) ||
                                name.Equals(resourceName, StringComparison.OrdinalIgnoreCase))
                            {
                                using (var altStream = assembly.GetManifestResourceStream(name))
                                {
                                    if (altStream != null)
                                        return WriteStreamToFile(altStream, targetPath, name);
                                }
                            }
                        }
                        
                        Helpers.Logger.Warn($"Embedded resource not found: {resourceName}");
                        return false;
                    }
                    
                    return WriteStreamToFile(stream, targetPath, resourceName);
                }
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error($"Failed to extract embedded resource: {resourceName}", ex);
                return false;
            }
        }

        private static bool EnsureEmbeddedResourceCurrent(string resourceName, string targetPath)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                        return WriteStreamToFileIfDifferent(stream, targetPath, resourceName);
                }

                foreach (var name in assembly.GetManifestResourceNames())
                {
                    if (!name.Equals(resourceName, StringComparison.OrdinalIgnoreCase)) continue;
                    using (var stream = assembly.GetManifestResourceStream(name))
                    {
                        if (stream != null)
                            return WriteStreamToFileIfDifferent(stream, targetPath, name);
                    }
                }

                Helpers.Logger.Warn($"Embedded resource not found: {resourceName}");
                return false;
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error($"Failed to refresh embedded resource: {resourceName}", ex);
                return false;
            }
        }

        private static bool WriteStreamToFile(Stream stream, string targetPath, string resourceName)
        {
            string targetDir = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrWhiteSpace(targetDir))
                Directory.CreateDirectory(targetDir);

            using (var fileStream = new FileStream(targetPath, FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(fileStream);
            }
            
            Helpers.Logger.Info($"Seeded from embedded resource: {resourceName}");
            return true;
        }

        private static bool WriteStreamToFileIfDifferent(Stream stream, string targetPath, string resourceName)
        {
            string targetDir = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrWhiteSpace(targetDir))
                Directory.CreateDirectory(targetDir);

            string tempPath = targetPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(fileStream);
                }

                if (FilesEqual(tempPath, targetPath))
                    return true;

                File.Copy(tempPath, targetPath, overwrite: true);
                Helpers.Logger.Info($"Updated from embedded resource: {resourceName}");
                return true;
            }
            finally
            {
                try
                {
                    if (File.Exists(tempPath)) File.Delete(tempPath);
                }
                catch { }
            }
        }

        private static bool FilesEqual(string firstPath, string secondPath)
        {
            if (!File.Exists(firstPath) || !File.Exists(secondPath)) return false;

            var firstInfo = new FileInfo(firstPath);
            var secondInfo = new FileInfo(secondPath);
            if (firstInfo.Length != secondInfo.Length) return false;

            const int BufferSize = 64 * 1024;
            var firstBuffer = new byte[BufferSize];
            var secondBuffer = new byte[BufferSize];

            using (var first = new FileStream(firstPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var second = new FileStream(secondPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                while (true)
                {
                    int firstRead = first.Read(firstBuffer, 0, firstBuffer.Length);
                    int secondRead = second.Read(secondBuffer, 0, secondBuffer.Length);
                    if (firstRead != secondRead) return false;
                    if (firstRead == 0) return true;

                    for (int i = 0; i < firstRead; i++)
                    {
                        if (firstBuffer[i] != secondBuffer[i]) return false;
                    }
                }
            }
        }
    }
}
