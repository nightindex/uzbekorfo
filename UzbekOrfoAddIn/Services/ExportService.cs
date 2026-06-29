using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using UzbekOrfoAddIn.Helpers;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Multi-format export service for spelling error reports.
    /// Supports TXT, CSV, HTML, JSON, and Word document export.
    /// </summary>
    public class ExportService
    {
        // =====================================================================
        //  EXPORT OPTIONS
        // =====================================================================

        /// <summary>
        /// Configurable export options set by the user in the export form.
        /// </summary>
        public class ExportOptions
        {
            public bool IncludeContext { get; set; } = true;
            public bool IncludeSuggestions { get; set; } = true;
            public bool IncludeParagraphNumbers { get; set; } = true;
            public bool IncludeStatistics { get; set; } = true;
            public bool IncludeTimestamp { get; set; } = true;
            public bool OnlyUnresolved { get; set; } = true;
            public string DocumentName { get; set; } = "";
        }

        // =====================================================================
        //  STATISTICS
        // =====================================================================

        /// <summary>
        /// Error report statistics computed from the error list.
        /// </summary>
        public class ErrorStatistics
        {
            public int TotalErrors { get; set; }
            public int ResolvedErrors { get; set; }
            public int UnresolvedErrors { get; set; }
            public int SpellingErrors { get; set; }
            public int GrammarErrors { get; set; }
            public int StyleErrors { get; set; }
            public int TypoErrors { get; set; }
            public Dictionary<string, int> MostCommonErrors { get; set; } = new Dictionary<string, int>();
            public int AffectedParagraphs { get; set; }
            public DateTime GeneratedAt { get; set; } = DateTime.Now;
        }

        /// <summary>
        /// Compute statistics from the error list.
        /// </summary>
        public static ErrorStatistics GetStatistics(List<ErrorEntry> errors)
        {
            if (errors == null || errors.Count == 0)
                return new ErrorStatistics();

            var stats = new ErrorStatistics
            {
                TotalErrors = errors.Count,
                ResolvedErrors = errors.Count(e => e.IsResolved),
                UnresolvedErrors = errors.Count(e => !e.IsResolved),
                SpellingErrors = errors.Count(e => e.Severity == ErrorSeverity.Spelling),
                GrammarErrors = errors.Count(e => e.Severity == ErrorSeverity.Grammar),
                StyleErrors = errors.Count(e => e.Severity == ErrorSeverity.Style),
                TypoErrors = errors.Count(e => e.Severity == ErrorSeverity.Typo),
                AffectedParagraphs = errors.Select(e => e.ParagraphIndex).Distinct().Count()
            };

            // Top 10 most repeated errors
            var grouped = errors
                .Where(e => !string.IsNullOrEmpty(e.Word))
                .GroupBy(e => e.Word.ToLowerInvariant())
                .OrderByDescending(g => g.Count())
                .Take(10);

            foreach (var g in grouped)
                stats.MostCommonErrors[g.Key] = g.Count();

            return stats;
        }

        // =====================================================================
        //  EXPORT TO TEXT
        // =====================================================================

        public static string ExportToText(List<ErrorEntry> errors, ExportOptions options)
        {
            var filtered = FilterErrors(errors, options);
            var sb = new StringBuilder();

            // Header
            sb.AppendLine("═══════════════════════════════════════════════════");
            sb.AppendLine("           ЎЗБЕК ОРФО — ХАТОЛАР ҲИСОБОТИ");
            sb.AppendLine("═══════════════════════════════════════════════════");
            sb.AppendLine();

            if (!string.IsNullOrEmpty(options.DocumentName))
                sb.AppendLine($"  Ҳужжат:    {options.DocumentName}");
            if (options.IncludeTimestamp)
                sb.AppendLine($"  Сана:      {DateTime.Now:dd.MM.yyyy  HH:mm}");
            sb.AppendLine($"  Хатолар:   {filtered.Count} та");
            sb.AppendLine();
            sb.AppendLine("───────────────────────────────────────────────────");
            sb.AppendLine();

            // Errors
            for (int i = 0; i < filtered.Count; i++)
            {
                var err = filtered[i];
                sb.Append($"  {i + 1,3}. {err.Word ?? "?"}");

                if (options.IncludeSuggestions && !string.IsNullOrEmpty(err.BestSuggestion))
                    sb.Append($"  →  {err.BestSuggestion}");

                if (options.IncludeParagraphNumbers)
                    sb.Append($"  (§{err.ParagraphIndex})");

                sb.AppendLine();

                if (options.IncludeContext && !string.IsNullOrEmpty(err.Context))
                {
                    sb.AppendLine($"       Контекст: {err.Context.Trim()}");
                }

                sb.AppendLine();
            }

            // Statistics footer
            if (options.IncludeStatistics)
            {
                var stats = GetStatistics(errors);
                sb.AppendLine("───────────────────────────────────────────────────");
                sb.AppendLine("  СТАТИСТИКА");
                sb.AppendLine($"    Жами хатолар:       {stats.TotalErrors}");
                sb.AppendLine($"    Тузатилган:         {stats.ResolvedErrors}");
                sb.AppendLine($"    Тузатилмаган:       {stats.UnresolvedErrors}");
                sb.AppendLine($"    Имло хатолари:      {stats.SpellingErrors}");
                sb.AppendLine($"    Грамматика:        {stats.GrammarErrors}");
                sb.AppendLine($"    Услуб хатолари:     {stats.StyleErrors}");
                sb.AppendLine($"    Тахминий хатолар:   {stats.TypoErrors}");
                sb.AppendLine($"    Параграфлар сони:   {stats.AffectedParagraphs}");

                if (stats.MostCommonErrors.Count > 0)
                {
                    sb.AppendLine();
                    sb.AppendLine("  Энг кўп хато:");
                    foreach (var kv in stats.MostCommonErrors)
                        sb.AppendLine($"    • {kv.Key}  ({kv.Value}x)");
                }
            }

            sb.AppendLine();
            sb.AppendLine("═══════════════════════════════════════════════════");
            sb.AppendLine("  Ўзбек Орфо | Автоматик ҳисобот");
            sb.AppendLine("═══════════════════════════════════════════════════");

            return sb.ToString();
        }

        // =====================================================================
        //  EXPORT TO CSV
        // =====================================================================

        public static string ExportToCsv(List<ErrorEntry> errors, ExportOptions options)
        {
            var filtered = FilterErrors(errors, options);
            var sb = new StringBuilder();

            // UTF-8 BOM for Excel compatibility
            sb.Append('\uFEFF');

            // Header row
            var headers = new List<string> { "№", "Сўз", "Тури", "Турлари" };
            if (options.IncludeSuggestions) headers.Add("Тузатиш варианти");
            if (options.IncludeParagraphNumbers) headers.Add("Параграф");
            if (options.IncludeContext) headers.Add("Контекст");
            headers.Add("Сана");
            sb.AppendLine(string.Join(",", headers));

            // Data rows
            for (int i = 0; i < filtered.Count; i++)
            {
                var err = filtered[i];
                var cells = new List<string>
                {
                    (i + 1).ToString(),
                    CsvEscape(err.Word ?? ""),
                    err.IsGrammarError ? "Грамматика" : "Имло",
                    err.Severity.ToString()
                };

                if (options.IncludeSuggestions)
                    cells.Add(CsvEscape(err.BestSuggestion ?? ""));
                if (options.IncludeParagraphNumbers)
                    cells.Add(err.ParagraphIndex.ToString());
                if (options.IncludeContext)
                    cells.Add(CsvEscape(err.Context ?? ""));
                cells.Add(err.DetectedAt.ToString("dd.MM.yyyy HH:mm"));

                sb.AppendLine(string.Join(",", cells));
            }

            return sb.ToString();
        }

        // =====================================================================
        //  EXPORT TO HTML
        // =====================================================================

        public static string ExportToHtml(List<ErrorEntry> errors, ExportOptions options)
        {
            var filtered = FilterErrors(errors, options);
            var stats = GetStatistics(errors);
            var sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"uz\">");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset=\"utf-8\">");
            sb.AppendLine($"<title>Хатолар ҳисоботи — {HtmlEscape(options.DocumentName)}</title>");
            sb.AppendLine("<style>");
            sb.AppendLine(@"
:root { --primary: #2563EB; --error: #DC2626; --success: #16A34A; --warning: #F59E0B; --bg: #FFFFFF; --surface: #F8FAFC; --text: #0F172A; --text2: #64748B; --border: #E2E8F0; }
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family: 'Segoe UI', Tahoma, sans-serif; background: var(--bg); color: var(--text); padding: 40px; max-width: 960px; margin: 0 auto; }
h1 { color: var(--primary); font-size: 24px; margin-bottom: 8px; }
.subtitle { color: var(--text2); font-size: 14px; margin-bottom: 24px; }
.stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin-bottom: 32px; }
.stat-card { background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 16px; }
.stat-card .number { font-size: 28px; font-weight: 700; color: var(--primary); }
.stat-card .label { font-size: 12px; color: var(--text2); margin-top: 4px; }
.stat-card.error .number { color: var(--error); }
.stat-card.success .number { color: var(--success); }
.stat-card.warning .number { color: var(--warning); }
table { width: 100%; border-collapse: collapse; margin-bottom: 24px; }
th { background: var(--primary); color: white; padding: 10px 14px; text-align: left; font-size: 13px; font-weight: 600; }
th:first-child { border-radius: 8px 0 0 0; }
th:last-child { border-radius: 0 8px 0 0; }
td { padding: 10px 14px; border-bottom: 1px solid var(--border); font-size: 13px; vertical-align: top; }
tr:hover { background: var(--surface); }
.word-error { color: var(--error); font-weight: 600; }
.word-fix { color: var(--success); }
.context { color: var(--text2); font-size: 12px; font-style: italic; }
.severity { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.severity.Spelling { background: #FEE2E2; color: #991B1B; }
.severity.Grammar { background: #DCFCE7; color: #166534; }
.severity.Style { background: #FEF3C7; color: #92400E; }
.severity.Typo { background: #DBEAFE; color: #1E40AF; }
.common-errors { margin-bottom: 24px; }
.common-errors h3 { font-size: 16px; margin-bottom: 12px; }
.bar-row { display: flex; align-items: center; margin-bottom: 6px; }
.bar-label { width: 120px; font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.bar-track { flex: 1; height: 20px; background: var(--surface); border-radius: 4px; overflow: hidden; }
.bar-fill { height: 100%; background: var(--primary); border-radius: 4px; min-width: 24px; display: flex; align-items: center; justify-content: flex-end; padding-right: 6px; color: white; font-size: 11px; font-weight: 600; }
footer { text-align: center; color: var(--text2); font-size: 12px; padding-top: 24px; border-top: 1px solid var(--border); }
");
            sb.AppendLine("</style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");

            // Title
            sb.AppendLine("<h1>📄 Хатолар ҳисоботи</h1>");
            sb.Append("<p class=\"subtitle\">");
            if (!string.IsNullOrEmpty(options.DocumentName))
                sb.Append($"Ҳужжат: <strong>{HtmlEscape(options.DocumentName)}</strong> &nbsp;|&nbsp; ");
            if (options.IncludeTimestamp)
                sb.Append($"Сана: {DateTime.Now:dd.MM.yyyy HH:mm}");
            sb.AppendLine("</p>");

            // Statistics cards
            if (options.IncludeStatistics)
            {
                sb.AppendLine("<div class=\"stats-grid\">");
                AppendStatCard(sb, stats.TotalErrors.ToString(), "Жами хатолар", "error");
                AppendStatCard(sb, stats.ResolvedErrors.ToString(), "Тузатилган", "success");
                AppendStatCard(sb, stats.UnresolvedErrors.ToString(), "Тузатилмаган", "warning");
                AppendStatCard(sb, stats.AffectedParagraphs.ToString(), "Параграфлар", "");
                sb.AppendLine("</div>");

                // Common errors bar chart
                if (stats.MostCommonErrors.Count > 0)
                {
                    int maxCount = stats.MostCommonErrors.Values.Max();
                    sb.AppendLine("<div class=\"common-errors\">");
                    sb.AppendLine("<h3>Энг кўп учраган хатолар</h3>");

                    foreach (var kv in stats.MostCommonErrors)
                    {
                        int pct = maxCount > 0 ? (int)(kv.Value * 100.0 / maxCount) : 0;
                        sb.AppendLine("<div class=\"bar-row\">");
                        sb.AppendLine($"  <span class=\"bar-label\">{HtmlEscape(kv.Key)}</span>");
                        sb.AppendLine($"  <div class=\"bar-track\"><div class=\"bar-fill\" style=\"width:{pct}%\">{kv.Value}</div></div>");
                        sb.AppendLine("</div>");
                    }

                    sb.AppendLine("</div>");
                }
            }

            // Error table
            sb.AppendLine("<table>");
            sb.Append("<tr><th>№</th><th>Сўз</th><th>Тури</th>");
            if (options.IncludeSuggestions) sb.Append("<th>Вариант</th>");
            if (options.IncludeParagraphNumbers) sb.Append("<th>§</th>");
            if (options.IncludeContext) sb.Append("<th>Контекст</th>");
            sb.AppendLine("</tr>");

            for (int i = 0; i < filtered.Count; i++)
            {
                var err = filtered[i];
                sb.Append("<tr>");
                sb.Append($"<td>{i + 1}</td>");
                sb.Append($"<td class=\"word-error\">{HtmlEscape(err.Word ?? "")}</td>");
                sb.Append($"<td><span class=\"severity {err.Severity}\">{SeverityLabel(err.Severity)}</span></td>");

                if (options.IncludeSuggestions)
                    sb.Append($"<td class=\"word-fix\">{HtmlEscape(err.BestSuggestion ?? "—")}</td>");
                if (options.IncludeParagraphNumbers)
                    sb.Append($"<td>{err.ParagraphIndex}</td>");
                if (options.IncludeContext)
                    sb.Append($"<td class=\"context\">{HtmlEscape(err.Context ?? "")}</td>");

                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</table>");

            // Footer
            sb.AppendLine($"<footer>Ўзбек Орфо — Автоматик ҳисобот | {DateTime.Now:dd.MM.yyyy HH:mm:ss}</footer>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb.ToString();
        }

        // =====================================================================
        //  EXPORT TO JSON
        // =====================================================================

        public static string ExportToJson(List<ErrorEntry> errors, ExportOptions options)
        {
            var filtered = FilterErrors(errors, options);
            var stats = GetStatistics(errors);
            var sb = new StringBuilder();

            sb.AppendLine("{");
            sb.AppendLine("  \"report\": {");
            sb.AppendLine($"    \"tool\": \"Ўзбек Орфо\",");
            sb.AppendLine($"    \"document\": {JsonEscape(options.DocumentName)},");
            sb.AppendLine($"    \"generatedAt\": \"{DateTime.Now:yyyy-MM-ddTHH:mm:ss}\",");
            sb.AppendLine($"    \"totalErrors\": {stats.TotalErrors},");
            sb.AppendLine($"    \"resolvedErrors\": {stats.ResolvedErrors},");
            sb.AppendLine($"    \"unresolvedErrors\": {stats.UnresolvedErrors},");
            sb.AppendLine($"    \"affectedParagraphs\": {stats.AffectedParagraphs}");
            sb.AppendLine("  },");

            // Statistics
            if (options.IncludeStatistics && stats.MostCommonErrors.Count > 0)
            {
                sb.AppendLine("  \"statistics\": {");
                sb.AppendLine($"    \"spelling\": {stats.SpellingErrors},");
                sb.AppendLine($"    \"grammar\": {stats.GrammarErrors},");
                sb.AppendLine($"    \"style\": {stats.StyleErrors},");
                sb.AppendLine($"    \"typo\": {stats.TypoErrors},");
                sb.AppendLine("    \"mostCommon\": [");

                var commonList = stats.MostCommonErrors.ToList();
                for (int j = 0; j < commonList.Count; j++)
                {
                    var comma = j < commonList.Count - 1 ? "," : "";
                    sb.AppendLine($"      {{ \"word\": {JsonEscape(commonList[j].Key)}, \"count\": {commonList[j].Value} }}{comma}");
                }

                sb.AppendLine("    ]");
                sb.AppendLine("  },");
            }

            // Errors array
            sb.AppendLine("  \"errors\": [");
            for (int i = 0; i < filtered.Count; i++)
            {
                var err = filtered[i];
                var comma = i < filtered.Count - 1 ? "," : "";
                sb.Append("    { ");
                sb.Append($"\"id\": {i + 1}");
                sb.Append($", \"word\": {JsonEscape(err.Word ?? "")}");
                sb.Append($", \"severity\": \"{err.Severity}\"");

                if (options.IncludeSuggestions)
                    sb.Append($", \"suggestion\": {JsonEscape(err.BestSuggestion ?? "")}");
                if (options.IncludeParagraphNumbers)
                    sb.Append($", \"paragraph\": {err.ParagraphIndex}");
                if (options.IncludeContext)
                    sb.Append($", \"context\": {JsonEscape(err.Context ?? "")}");

                if (err.IsGrammarError)
                {
                    sb.Append($", \"ruleId\": {JsonEscape(err.RuleId ?? "")}");
                    sb.Append($", \"message\": {JsonEscape(err.Message ?? "")}");
                    sb.Append($", \"category\": {JsonEscape(err.Category ?? "")}");
                }

                sb.Append($", \"resolved\": {(err.IsResolved ? "true" : "false")}");
                sb.Append($", \"detectedAt\": \"{err.DetectedAt:yyyy-MM-ddTHH:mm:ss}\"");
                sb.AppendLine($" }}{comma}");
            }

            sb.AppendLine("  ]");
            sb.AppendLine("}");

            return sb.ToString();
        }

        // =====================================================================
        //  GENERATE PREVIEW (truncated for UI)
        // =====================================================================

        /// <summary>
        /// Generates a truncated preview of the export output for the preview panel.
        /// </summary>
        public static string GeneratePreview(List<ErrorEntry> errors, ExportFormat format, ExportOptions options, int maxLines = 30)
        {
            string full;
            switch (format)
            {
                case ExportFormat.Text: full = ExportToText(errors, options); break;
                case ExportFormat.Csv: full = ExportToCsv(errors, options); break;
                case ExportFormat.Html: full = ExportToHtml(errors, options); break;
                case ExportFormat.Json: full = ExportToJson(errors, options); break;
                default: full = ExportToText(errors, options); break;
            }

            var lines = full.Split(new[] { '\n' }, StringSplitOptions.None);
            if (lines.Length <= maxLines)
                return full;

            var preview = new StringBuilder();
            for (int i = 0; i < maxLines; i++)
                preview.AppendLine(lines[i]);
            preview.AppendLine($"... ({lines.Length - maxLines} қатор яширилган)");

            return preview.ToString();
        }

        // =====================================================================
        //  SAVE TO FILE
        // =====================================================================

        /// <summary>
        /// Save the exported content to a file.
        /// </summary>
        public static void SaveToFile(string filePath, List<ErrorEntry> errors, ExportFormat format, ExportOptions options)
        {
            string content;
            switch (format)
            {
                case ExportFormat.Text: content = ExportToText(errors, options); break;
                case ExportFormat.Csv: content = ExportToCsv(errors, options); break;
                case ExportFormat.Html: content = ExportToHtml(errors, options); break;
                case ExportFormat.Json: content = ExportToJson(errors, options); break;
                default: content = ExportToText(errors, options); break;
            }

            File.WriteAllText(filePath, content, Encoding.UTF8);
            Logger.Info($"Экспорт сақланди: {filePath} ({format})");
        }

        /// <summary>
        /// Returns the appropriate file extension for the given format.
        /// </summary>
        public static string GetExtension(ExportFormat format)
        {
            switch (format)
            {
                case ExportFormat.Text: return ".txt";
                case ExportFormat.Csv: return ".csv";
                case ExportFormat.Html: return ".html";
                case ExportFormat.Json: return ".json";
                default: return ".txt";
            }
        }

        /// <summary>
        /// Returns a user-friendly label for a format.
        /// </summary>
        public static string GetFormatLabel(ExportFormat format)
        {
            switch (format)
            {
                case ExportFormat.Text: return "Матн файл (.txt)";
                case ExportFormat.Csv: return "CSV жадвал (.csv)";
                case ExportFormat.Html: return "HTML ҳисобот (.html)";
                case ExportFormat.Json: return "JSON маълумот (.json)";
                case ExportFormat.Word: return "Word ҳужжат (.docx)";
                default: return format.ToString();
            }
        }

        // =====================================================================
        //  HELPERS
        // =====================================================================

        private static List<ErrorEntry> FilterErrors(List<ErrorEntry> errors, ExportOptions options)
        {
            if (errors == null) return new List<ErrorEntry>();
            return options.OnlyUnresolved
                ? errors.Where(e => !e.IsResolved).ToList()
                : errors.ToList();
        }

        private static string CsvEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return "\"\"";
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static string HtmlEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;");
        }

        private static string JsonEscape(string value)
        {
            if (value == null) return "null";
            return "\"" + value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t") + "\"";
        }

        private static string SeverityLabel(ErrorSeverity severity)
        {
            switch (severity)
            {
                case ErrorSeverity.Spelling: return "Имло";
                case ErrorSeverity.Grammar: return "Грамматика";
                case ErrorSeverity.Style: return "Услуб";
                case ErrorSeverity.Typo: return "Хато";
                default: return severity.ToString();
            }
        }

        private static void AppendStatCard(StringBuilder sb, string number, string label, string cssClass)
        {
            sb.AppendLine($"<div class=\"stat-card {cssClass}\">");
            sb.AppendLine($"  <div class=\"number\">{number}</div>");
            sb.AppendLine($"  <div class=\"label\">{label}</div>");
            sb.AppendLine("</div>");
        }
    }
}
