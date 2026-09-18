using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Prediction
{
    /// <summary>Display-only official-name casing; never changes text already typed or spelling validity.</summary>
    public static class PredictionCasing
    {
        private static readonly string[][] Names = Load();

        private static string[][] Load()
        {
            using (var stream = typeof(PredictionCasing).Assembly.GetManifestResourceStream("UzbekOrfoAddIn.Data.prediction_casing.json"))
            {
                if (stream == null) throw new InvalidDataException("Missing prediction casing data.");
                using (var reader = new StreamReader(stream))
                {
                    var names = JsonConvert.DeserializeObject<string[]>(reader.ReadToEnd());
                    if (names == null || names.Any(n => string.IsNullOrWhiteSpace(n) || n.Split(' ').Length < 2))
                        throw new InvalidDataException("Invalid prediction casing data.");
                    return names.Select(n => n.Split(' ')).ToArray();
                }
            }
        }

        public static string Apply(string context, string prefix, string completion)
        {
            if (completion == null || prefix == null || !completion.StartsWith(prefix, StringComparison.Ordinal)) return completion;
            context = context ?? string.Empty;
            int boundary = context.LastIndexOfAny(new[] { '.', '!', '?', ';', ':', '\r', '\n', '\t', '\a' });
            var preceding = TextHelper.Tokenize(context.Substring(boundary + 1)).Select(t => t.Original).ToArray();
            var display = completion.Split(' ');
            var words = preceding.Concat(display).Select(TextHelper.NormalizeWord).ToArray();
            foreach (var name in Names)
            {
                var key = name.Select(TextHelper.NormalizeWord).ToArray();
                for (int start = Math.Max(0, preceding.Length - name.Length + 1); start + name.Length <= words.Length; start++)
                {
                    bool matches = true;
                    for (int offset = 0; offset < name.Length; offset++)
                    {
                        // Retain observed suffix spelling on the last word (Республикасининг, etc.).
                        if (offset == name.Length - 1 ? !words[start + offset].StartsWith(key[offset], StringComparison.Ordinal)
                            : words[start + offset] != key[offset]) { matches = false; break; }
                    }
                    if (!matches) continue;
                    for (int offset = 0; offset < name.Length; offset++)
                    {
                        int index = start + offset - preceding.Length;
                        if (index >= 0 && char.IsUpper(name[offset][0]))
                            display[index] = char.ToUpperInvariant(display[index][0]) + display[index].Substring(1);
                    }
                }
            }
            string result = string.Join(" ", display);
            return prefix + result.Substring(prefix.Length);
        }
    }
}
