using System;
using System.IO;
using System.Reflection;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Copies bundled seed data files from the add-in folder into the runtime
    /// AppData directory when user-local files are missing.
    /// Supports both file-based (F5 debug) and embedded resource (ClickOnce) deployment.
    /// </summary>
    public static class DataSeedService
    {
        public static void SeedDictionary(string mainDictionaryPath)
        {
            if (string.IsNullOrWhiteSpace(mainDictionaryPath)) return;
            try
            {
                if (File.Exists(mainDictionaryPath)) return;
                
                // Try file-based first (works in F5 debug), then embedded resource (ClickOnce)
                if (!CopyBundledDataFile("uzbek_main.dic", mainDictionaryPath, overwrite: false))
                {
                    ExtractEmbeddedResource("UzbekOrfoAddIn.Data.uzbek_main.dic", mainDictionaryPath);
                }
            }
            catch (Exception ex)
            {
                Helpers.Logger.Error("Failed to seed dictionary", ex);
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
    }
}
