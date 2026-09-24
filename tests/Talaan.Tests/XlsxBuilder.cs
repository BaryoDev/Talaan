using System.IO.Compression;
using System.Text;

namespace Talaan.Tests;

/// <summary>
/// Builds a minimal .xlsx (a zip of OOXML parts) in memory, so xlsx tests can assert against exact
/// bytes instead of a binary fixture. Every part is a raw string handed straight to the zip entry:
/// nothing here escapes or validates it, so a test can write malformed or unusual XML (a DOCTYPE, an
/// out-of-range style index, an oversized column reference) on purpose.
/// </summary>
internal static class XlsxBuilder
{
    private const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string DefaultWorkbookXml = """
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
        """;

    private const string DefaultWorkbookRelsXml = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        </Relationships>
        """;

    /// <summary>
    /// Zips up a workbook from its parts. Only <paramref name="sheetXml"/> is required; the rest
    /// default to a plain single-sheet workbook with no shared strings or styles. Pass
    /// <paramref name="workbookXml"/>, <paramref name="workbookRelsXml"/> and <paramref name="sheetPath"/>
    /// together to control how the first sheet resolves through the rels file.
    /// <paramref name="extraParts"/> adds further raw parts verbatim (path to content), for a decoy
    /// worksheet, a second sheet, or anything else a test needs beyond the parts above.
    /// </summary>
    public static MemoryStream Build(
        string sheetXml,
        string? sharedStringsXml = null,
        string? stylesXml = null,
        string? workbookXml = null,
        string? workbookRelsXml = null,
        string sheetPath = "xl/worksheets/sheet1.xml",
        IReadOnlyDictionary<string, string>? extraParts = null)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string path, string content)
            {
                using var w = new StreamWriter(zip.CreateEntry(path).Open(), new UTF8Encoding(false));
                w.Write(content);
            }

            Add("xl/workbook.xml", workbookXml ?? DefaultWorkbookXml);
            Add("xl/_rels/workbook.xml.rels", workbookRelsXml ?? DefaultWorkbookRelsXml);
            Add(sheetPath, sheetXml);
            if (sharedStringsXml != null) Add("xl/sharedStrings.xml", sharedStringsXml);
            if (stylesXml != null) Add("xl/styles.xml", stylesXml);
            if (extraParts != null)
                foreach (var (path, content) in extraParts) Add(path, content);
        }
        ms.Position = 0;
        return ms;
    }

    /// <summary>A single-row worksheet, row 1, holding <paramref name="cellsXml"/> verbatim.</summary>
    public static string Row(string cellsXml) => Sheet(RowAt(1, cellsXml));

    /// <summary>A `&lt;row&gt;` element at the given (1-based) row reference, for tests that need to
    /// skip or reorder rows (see issue #7).</summary>
    public static string RowAt(int r, string cellsXml) => $"""<row r="{r}">{cellsXml}</row>""";

    /// <summary>Wraps already-built `&lt;row&gt;` elements (e.g. from <see cref="RowAt"/>) in a worksheet.</summary>
    public static string Sheet(params string[] rowsXml) =>
        $"""<worksheet xmlns="{Ns}"><sheetData>{string.Concat(rowsXml)}</sheetData></worksheet>""";
}
