using System;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

namespace UzbekOrfoAddIn.Helpers
{
    /// <summary>
    /// Shared helper for writing lightweight .xlsx files (Open XML)
    /// without depending on any third-party library.
    /// Used by both error-export and dictionary-migration flows.
    /// </summary>
    internal static class XlsxExportHelper
    {
        // ───────────────────── XLSX package boilerplate ─────────────────────

        /// <summary>
        /// Writes a complete single-sheet .xlsx file.
        /// The caller supplies only the sheet name and the sheet XML content.
        /// </summary>
        public static void WriteXlsxPackage(string filePath, string sheetName, string sheetXml)
        {
            using (var fs = File.Create(filePath))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create, false))
            {
                WriteZipEntry(zip, "[Content_Types].xml",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
  <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
  <Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/>
  <Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml""/>
</Types>");

                WriteZipEntry(zip, "_rels/.rels",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/>
  <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml""/>
</Relationships>");

                WriteZipEntry(zip, "docProps/app.xml",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties""
            xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">
  <Application>UzbekOrfoAddIn</Application>
</Properties>");

                var createdUtc = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                WriteZipEntry(zip, "docProps/core.xml",
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties""
                   xmlns:dc=""http://purl.org/dc/elements/1.1/""
                   xmlns:dcterms=""http://purl.org/dc/terms/""
                   xmlns:dcmitype=""http://purl.org/dc/dcmitype/""
                   xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <dc:creator>UzbekOrfoAddIn</dc:creator>
  <cp:lastModifiedBy>UzbekOrfoAddIn</cp:lastModifiedBy>
  <dcterms:created xsi:type=""dcterms:W3CDTF"">{createdUtc}</dcterms:created>
  <dcterms:modified xsi:type=""dcterms:W3CDTF"">{createdUtc}</dcterms:modified>
</cp:coreProperties>");

                string escapedSheetName = EscapeXmlForXlsx(sheetName);
                WriteZipEntry(zip, "xl/workbook.xml",
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
          xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <sheets>
    <sheet name=""{escapedSheetName}"" sheetId=""1"" r:id=""rId1""/>
  </sheets>
</workbook>");

                WriteZipEntry(zip, "xl/_rels/workbook.xml.rels",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>
</Relationships>");

                WriteZipEntry(zip, "xl/styles.xml",
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <fonts count=""1""><font><sz val=""11""/><name val=""Calibri""/></font></fonts>
  <fills count=""1""><fill><patternFill patternType=""none""/></fill></fills>
  <borders count=""1""><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>
  <cellXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/></cellXfs>
  <cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>
</styleSheet>");

                WriteZipEntry(zip, "xl/worksheets/sheet1.xml", sheetXml);
            }
        }

        // ───────────────────── XLS conversion via COM ─────────────────────

        /// <summary>
        /// Converts a .xlsx file to .xls (Excel 97-2003) by launching Excel via COM.
        /// Throws <see cref="InvalidOperationException"/> if Excel is not installed.
        /// </summary>
        public static void ConvertXlsxToXls(string sourceXlsx, string targetXls)
        {
            object excelApp = null;
            object workbooks = null;
            object workbook = null;

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                    throw new InvalidOperationException("XLS экспорт учун Excel ўрнатилмаган.");

                excelApp = Activator.CreateInstance(excelType);
                dynamic app = excelApp;
                app.Visible = false;
                app.DisplayAlerts = false;

                workbooks = app.Workbooks;
                dynamic books = workbooks;
                workbook = books.Open(sourceXlsx);
                dynamic wb = workbook;

                // 56 = xlExcel8 (Excel 97-2003 Workbook .xls)
                wb.SaveAs(targetXls, 56);
                wb.Close(false);
                app.Quit();
            }
            finally
            {
                try { if (workbook != null) ((dynamic)workbook).Close(false); } catch { }
                try { if (excelApp != null) ((dynamic)excelApp).Quit(); } catch { }
                ReleaseComObjectSafe(workbook);
                ReleaseComObjectSafe(workbooks);
                ReleaseComObjectSafe(excelApp);
            }
        }

        // ───────────────────── Sheet-XML primitives ─────────────────────

        /// <summary>
        /// Appends a single row of inline-string cells to <paramref name="sb"/>.
        /// </summary>
        public static void AppendInlineRow(StringBuilder sb, int rowIndex, string[] values)
        {
            sb.Append("    <row r=\"").Append(rowIndex).AppendLine("\">");
            for (int i = 0; i < values.Length; i++)
            {
                string cellRef = GetExcelColumnName(i + 1) + rowIndex;
                string value = EscapeXmlForXlsx(values[i] ?? string.Empty);
                sb.Append("      <c r=\"").Append(cellRef).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">")
                  .Append(value)
                  .AppendLine("</t></is></c>");
            }
            sb.AppendLine("    </row>");
        }

        /// <summary>
        /// Converts a 1-based column number to an Excel column name (1 → A, 27 → AA).
        /// </summary>
        public static string GetExcelColumnName(int columnNumber)
        {
            var sb = new StringBuilder();
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                sb.Insert(0, (char)('A' + modulo));
                columnNumber = (columnNumber - modulo) / 26;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Escapes a string value for safe inclusion in an xlsx XML cell,
        /// stripping invalid XML characters.
        /// </summary>
        public static string EscapeXmlForXlsx(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;

            var sb = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++)
            {
                char ch = value[i];
                if (!IsValidXmlChar(ch)) continue;

                switch (ch)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    case '\'': sb.Append("&apos;"); break;
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }

        // ───────────────────── Internal helpers ─────────────────────

        private static void WriteZipEntry(ZipArchive zip, string entryPath, string content)
        {
            var entry = zip.CreateEntry(entryPath, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(content);
            }
        }

        private static bool IsValidXmlChar(int c)
        {
            return c == 0x9 || c == 0xA || c == 0xD ||
                   (c >= 0x20 && c <= 0xD7FF) ||
                   (c >= 0xE000 && c <= 0xFFFD);
        }

        private static void ReleaseComObjectSafe(object comObject)
        {
            if (comObject == null) return;
            try
            {
                if (Marshal.IsComObject(comObject))
                    Marshal.FinalReleaseComObject(comObject);
            }
            catch { }
        }
    }
}
