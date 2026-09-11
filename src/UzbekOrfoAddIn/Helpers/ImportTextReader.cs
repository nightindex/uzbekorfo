using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Reads text-based dictionary imports without corrupting Uzbek Cyrillic
    /// characters and parses CSV/TSV records with quoted fields.
    /// </summary>
    public static class ImportTextReader
    {
        private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);

        /// <summary>
        /// Reads UTF-8 (with or without a BOM), UTF-16 BOM files, and legacy
        /// Windows-1251 Cyrillic files. Invalid UTF-8 falls back to Windows-1251,
        /// the common legacy encoding for Uzbek Cyrillic word lists.
        /// </summary>
        public static string ReadTextFile(string filePath)
        {
            try
            {
                using (var stream = File.OpenRead(filePath))
                using (var reader = new StreamReader(stream, StrictUtf8, true))
                    return reader.ReadToEnd();
            }
            catch (DecoderFallbackException)
            {
                return File.ReadAllText(filePath, Encoding.GetEncoding(1251));
            }
        }

        /// <summary>
        /// Parses a delimited text file while respecting escaped quotes, quoted
        /// delimiters, and newline characters inside quoted fields.
        /// </summary>
        public static List<string[]> ParseDelimitedRecords(string text)
        {
            return ParseDelimitedRecords(text, DetectDelimiter(text));
        }

        private static List<string[]> ParseDelimitedRecords(string text, char delimiter)
        {
            var records = new List<string[]>();
            if (string.IsNullOrEmpty(text)) return records;

            var fields = new List<string>();
            var value = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < text.Length && text[i + 1] == '"')
                    {
                        value.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                    continue;
                }

                if (c == delimiter && !inQuotes)
                {
                    AddField(fields, value);
                    continue;
                }

                if ((c == '\r' || c == '\n') && !inQuotes)
                {
                    AddField(fields, value);
                    AddRecord(records, fields);
                    if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    continue;
                }

                value.Append(c);
            }

            if (value.Length > 0 || fields.Count > 0)
            {
                AddField(fields, value);
                AddRecord(records, fields);
            }

            return records;
        }

        private static void AddField(List<string> fields, StringBuilder value)
        {
            fields.Add(value.ToString().Trim());
            value.Clear();
        }

        private static void AddRecord(List<string[]> records, List<string> fields)
        {
            if (fields.Count == 1 && string.IsNullOrWhiteSpace(fields[0]))
            {
                fields.Clear();
                return;
            }

            records.Add(fields.ToArray());
            fields.Clear();
        }

        private static char DetectDelimiter(string text)
        {
            if (string.IsNullOrEmpty(text)) return ',';

            int comma = 0, semicolon = 0, tab = 0;
            bool inQuotes = false;
            bool hasContent = false;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < text.Length && text[i + 1] == '"')
                    {
                        i++;
                        hasContent = true;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                    continue;
                }

                if ((c == '\r' || c == '\n') && !inQuotes)
                {
                    if (hasContent) break;
                    continue;
                }

                if (!inQuotes)
                {
                    if (c == ',')
                    {
                        comma++;
                        hasContent = true;
                        continue;
                    }
                    if (c == ';')
                    {
                        semicolon++;
                        hasContent = true;
                        continue;
                    }
                    if (c == '\t')
                    {
                        tab++;
                        hasContent = true;
                        continue;
                    }
                }

                if (char.IsWhiteSpace(c)) continue;
                hasContent = true;
            }

            if (tab >= semicolon && tab >= comma && tab > 0) return '\t';
            if (semicolon >= comma && semicolon > 0) return ';';
            return ',';
        }
    }
}
