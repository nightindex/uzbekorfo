using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Xml;
using HtmlAgilityPack;
using NPOI.HWPF;
using UzbekOrfoAddIn.Helpers;

namespace UzbekOrfoAddIn.Prediction
{
    /// <summary>Read-only extraction; never starts Word or writes a source document.</summary>
    public sealed class DocumentExtractor
    {
        static DocumentExtractor()
        {
#if NET8_0_OR_GREATER
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
        }
        public const long MaxFileBytes = 50L * 1024 * 1024;
        public const int MaxExtractedCharacters = 4 * 1024 * 1024;
        public const int MaxPassages = 20000;
        private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string StrictWordNamespace = "http://purl.oclc.org/ooxml/wordprocessingml/main";

        public ExtractedDocument Extract(string path, CancellationToken cancellation = default(CancellationToken))
        {
            cancellation.ThrowIfCancellationRequested();
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (input.Length > MaxFileBytes) throw new InvalidDataException("File exceeds the 50 MiB limit.");
                string extension = System.IO.Path.GetExtension(path).ToLowerInvariant();
                var output = new PassageWriter(cancellation);
                if (extension == ".docx") ExtractDocx(input, output, cancellation);
                else if (extension == ".txt") output.Append(ImportTextReader.ReadTextFile(path));
                else if (extension == ".doc" || extension == ".html" || extension == ".htm")
                {
                    var signature = new byte[8];
                    int count = input.Read(signature, 0, signature.Length);
                    input.Position = 0;
                    if (count == 8 && BitConverter.ToString(signature) == "D0-CF-11-E0-A1-B1-1A-E1")
                        ExtractBinaryDoc(input, output, cancellation);
                    else
                    {
                        string html = ImportTextReader.ReadTextFile(path);
                        if (html.Length > MaxExtractedCharacters * 2)
                            throw new InvalidDataException("HTML exceeds the 8 Mi-character parsing limit.");
                        if (html.IndexOf('<') < 0 || html.TrimStart().StartsWith("{\\rtf", StringComparison.OrdinalIgnoreCase))
                            throw new NotSupportedException("This DOC is neither HTML nor a supported Word binary document.");
                        var document = new HtmlDocument();
                        document.LoadHtml(html);
                        ExtractHtml(document.DocumentNode.SelectSingleNode("//body") ?? document.DocumentNode, output, cancellation, 0);
                    }
                }
                else throw new NotSupportedException("Supported files: DOCX, DOC (HTML or Word binary), TXT, HTML and HTM.");
                output.Break();
                cancellation.ThrowIfCancellationRequested();
                return output.Document;
            }
        }

        private static void ExtractDocx(Stream input, PassageWriter output, CancellationToken cancellation)
        {
            using (var archive = new ZipArchive(input, ZipArchiveMode.Read, true))
            {
                var entry = archive.GetEntry("word/document.xml");
                if (entry == null) throw new InvalidDataException("DOCX has no word/document.xml.");
                const int maxXml = 32 * 1024 * 1024;
                if (entry.Length > maxXml) throw new InvalidDataException("DOCX body XML exceeds the 32 MiB limit.");
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                    MaxCharactersInDocument = maxXml, IgnoreComments = true
                };
                using (var stream = entry.Open())
                using (var reader = XmlReader.Create(stream, settings))
                {
                    int bodyDepth = -1;
                    int textDepth = -1;
                    bool foundBody = false;
                    bool available = reader.Read();
                    while (available)
                    {
                        cancellation.ThrowIfCancellationRequested();
                        bool word = reader.NamespaceURI == WordNamespace || reader.NamespaceURI == StrictWordNamespace;
                        if (word && reader.NodeType == XmlNodeType.Element)
                        {
                            string name = reader.LocalName;
                            if (name == "body") { bodyDepth = reader.Depth; foundBody = true; }
                            else if (bodyDepth >= 0)
                            {
                                if (name == "del" || name == "moveFrom" || name == "comment" || name == "comments" ||
                                    name == "hdr" || name == "ftr" || name == "instrText" || name == "delText")
                                {
                                    reader.Skip();
                                    available = !reader.EOF;
                                    continue;
                                }
                                if (name == "t" && !reader.IsEmptyElement) textDepth = reader.Depth;
                                else if (name == "tab") output.Append(" ");
                                else if (name == "br" || name == "cr" || name == "p") output.Break();
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement && word)
                        {
                            if (reader.LocalName == "t") textDepth = -1;
                            if (reader.LocalName == "p" || reader.LocalName == "tc") output.Break();
                            if (reader.LocalName == "body") bodyDepth = -1;
                        }
                        else if (bodyDepth >= 0 && textDepth >= 0 &&
                            (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.SignificantWhitespace ||
                             reader.NodeType == XmlNodeType.Whitespace || reader.NodeType == XmlNodeType.CDATA))
                            output.Append(reader.Value);
                        available = reader.Read();
                    }
                    if (!foundBody) throw new InvalidDataException("DOCX has no Word body.");
                }
            }
        }

        private static void ExtractHtml(HtmlNode node, PassageWriter output, CancellationToken cancellation, int depth)
        {
            cancellation.ThrowIfCancellationRequested();
            if (depth > 256) throw new InvalidDataException("HTML nesting exceeds 256 levels.");
            if (node.NodeType == HtmlNodeType.Comment) return;
            string name = node.Name.ToLowerInvariant();
            string style = node.GetAttributeValue("style", "").Replace(" ", "").ToLowerInvariant();
            string classes = node.GetAttributeValue("class", "").ToLowerInvariant();
            if (name == "script" || name == "style" || name == "head" || name == "del" || name == "noscript" ||
                name == "header" || name == "footer" || name == "nav" || node.Attributes["hidden"] != null ||
                style.Contains("display:none") || style.Contains("mso-element:comment") ||
                style.Contains("mso-element:header") || style.Contains("mso-element:footer") ||
                classes.Contains("msocomment") || classes.Contains("msocomtxt") || classes.Contains("msocomtext") ||
                classes.Contains("msodeletedtext")) return;
            // LexUZ export metadata and editorial source notes are not drafting prose.
            var classNames = classes.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string css in classNames)
                if (css == "comment" || css == "comment_for_warning" || css == "act_info" ||
                    css == "act_source" || css == "official_sour_text") return;
            if (node.NodeType == HtmlNodeType.Text)
            {
                output.Append(HtmlEntity.DeEntitize(((HtmlTextNode)node).Text));
                return;
            }
            bool block = name == "p" || name == "div" || name == "br" || name == "li" || name == "tr" ||
                name == "td" || name == "th" || name == "hr" || name == "section" || name == "article" ||
                name == "blockquote" || name == "pre" || (name.Length == 2 && name[0] == 'h' && name[1] >= '1' && name[1] <= '6');
            if (block) output.Break();
            foreach (var child in node.ChildNodes) ExtractHtml(child, output, cancellation, depth + 1);
            if (block) output.Break();
        }

        private static void ExtractBinaryDoc(Stream input, PassageWriter output, CancellationToken cancellation)
        {
            // GetRange is MAIN only: comments, headers, footers and footnotes are separate stories.
            var document = new HWPFDocument(input);
            {
                var range = document.GetRange();
                var fields = new Stack<bool>();
                for (int index = 0; index < range.NumCharacterRuns; index++)
                {
                    cancellation.ThrowIfCancellationRequested();
                    var run = range.GetCharacterRun(index);
                    if (run.IsMarkedDeleted() || run.IsVanished()) continue;
                    foreach (char character in run.Text)
                    {
                        if (character == '\u0013') { fields.Push(false); continue; }
                        if (character == '\u0014') { if (fields.Count > 0) { fields.Pop(); fields.Push(true); } continue; }
                        if (character == '\u0015') { if (fields.Count > 0) fields.Pop(); continue; }
                        if (fields.Contains(false)) continue;
                        if (character == '\r' || character == '\n' || character == '\u0007' || character == '\u000b') output.Break();
                        else if (!char.IsControl(character) || character == '\t') output.Append(character.ToString());
                    }
                }
            }
        }

        private sealed class PassageWriter
        {
            internal readonly ExtractedDocument Document = new ExtractedDocument();
            private readonly StringBuilder buffer = new StringBuilder();
            private readonly CancellationToken cancellation;
            private int characters;
            internal PassageWriter(CancellationToken cancellation) { this.cancellation = cancellation; }
            internal void Append(string text)
            {
                if (text.Length > MaxExtractedCharacters - characters)
                    throw new InvalidDataException("Extracted text exceeds the 4 Mi-character limit.");
                characters += text.Length;
                foreach (char character in text)
                {
                    if ((buffer.Length & 4095) == 0) cancellation.ThrowIfCancellationRequested();
                    if (character == '\r' || character == '\n') Break();
                    else buffer.Append(character);
                }
            }
            internal void Break()
            {
                string text = buffer.ToString().Trim();
                buffer.Clear();
                if (text.Length == 0) return;
                if (Document.Passages.Count == MaxPassages) throw new InvalidDataException("Document exceeds 20,000 passages.");
                Document.Passages.Add(text);
            }
        }
    }
}
