using System;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Creates and owns the runtime service graph used by the add-in.
    /// </summary>
    public sealed class AddInRuntime
    {
        public SettingsManager Settings { get; private set; }
        public ErrorStore ErrorStore { get; private set; }
        public DictionaryService DictionaryService { get; private set; }
        public SpellingEngine SpellingEngine { get; private set; }
        public AutoCorrectService AutoCorrectService { get; private set; }
        public TransliterationService Transliterator { get; private set; }
        public ExplanationProvider ExplanationProvider { get; private set; }
        public UzbekMorphAnalyzer MorphAnalyzer { get; private set; }
        public GrammarEngine GrammarEngine { get; private set; }

        private AddInRuntime() { }

        public static AddInRuntime Initialize(Word.Application application)
        {
            if (application == null) throw new ArgumentNullException(nameof(application));

            var runtime = new AddInRuntime
            {
                Settings = new SettingsManager(),
                ErrorStore = new ErrorStore()
            };

            DataSeedService.SeedDictionary(runtime.Settings.MainDictionaryPath);

            runtime.DictionaryService = new DictionaryService(
                runtime.Settings.MainDictionaryPath,
                runtime.Settings.UserDictionaryPath);
            runtime.DictionaryService.Load();
            Logger.Info($"Луғат юкланди: {runtime.DictionaryService.TotalWordCount} сўз");

            runtime.SpellingEngine = new SpellingEngine(runtime.DictionaryService);

            runtime.AutoCorrectService = new AutoCorrectService(runtime.SpellingEngine, runtime.DictionaryService);
            runtime.AutoCorrectService.Initialize(application);
            runtime.AutoCorrectService.IsEnabled = runtime.Settings.AutoCorrectEnabled;

            runtime.Transliterator = new TransliterationService(runtime.Settings.ExceptionsPath);
            DataSeedService.SeedTranslitExceptions(runtime.Settings.ExceptionsPath, runtime.Transliterator);
            Logger.Info($"Транслитерация юкланди: {runtime.Transliterator.GetExceptions().Count} та истисно");

            runtime.DictionaryService.SetTransliterator(runtime.Transliterator);

            runtime.ExplanationProvider = new ExplanationProvider(runtime.Settings.ExplanationsPath);
            DataSeedService.SeedExplanations(runtime.Settings.ExplanationsPath);
            runtime.ExplanationProvider.Load();
            Logger.Info($"Изоҳлар юкланди: {runtime.ExplanationProvider.EntryCount} та сўз");

            runtime.MorphAnalyzer = new UzbekMorphAnalyzer(runtime.DictionaryService, runtime.Transliterator);
            DataSeedService.SeedGrammarFiles(
                runtime.Settings.SuffixesPath,
                runtime.Settings.GrammarRulesPath,
                runtime.Settings.ProperNounsPath);
            runtime.MorphAnalyzer.LoadSuffixes(runtime.Settings.SuffixesPath);
            Logger.Info("Морфологик анализатор юкланди");

            runtime.GrammarEngine = new GrammarEngine(
                runtime.DictionaryService,
                runtime.Transliterator,
                runtime.MorphAnalyzer,
                runtime.Settings.GrammarRulesPath,
                runtime.Settings.ProperNounsPath);
            runtime.GrammarEngine.LoadRules();
            runtime.GrammarEngine.IsEnabled = runtime.Settings.GrammarCheckEnabled;
            Logger.Info($"Грамматика текшируви юкланди: {runtime.GrammarEngine.GetRules().Count} та қоида");

            Task.Run(() =>
            {
                try
                {
                    runtime.DictionaryService.RebuildSearchIndexBackground();
                    runtime.DictionaryService.WarmUpCaches();
                }
                catch (Exception ex) { Logger.Error("Cache warm-up failed (non-fatal)", ex); }
            });

            Logger.Info($"Инициализация муваффақиятли. Маълумотлар: {runtime.Settings.DataDirectory}");
            return runtime;
        }

        public void Shutdown()
        {
            try { AutoCorrectService?.Dispose(); }
            catch (Exception ex) { Logger.Warn($"AutoCorrect dispose error: {ex.Message}"); }

            try { DictionaryService?.Save(); }
            catch (Exception ex) { Logger.Warn($"Dictionary save error: {ex.Message}"); }

            try { Transliterator?.SaveExceptions(); }
            catch (Exception ex) { Logger.Warn($"Transliteration save error: {ex.Message}"); }

            try { Settings?.Save(); }
            catch (Exception ex) { Logger.Warn($"Settings save error: {ex.Message}"); }
        }
    }
}
