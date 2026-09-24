using System.Runtime.CompilerServices;
using System.Xml;
using Talaan;
using Talaan.Xlsx;
using Xunit;

namespace Talaan.Tests;

/// <summary>
/// Covers issue #11: no XML part Talaan reads should process a DTD. A DOCTYPE is not something a
/// legitimate .xlsx ever contains, so every part with one is expected to throw, not to expand
/// entities or to silently drop the offending value.
/// </summary>
public class XlsxSecurityTests
{
    // Four levels of nesting, base entity of 10 chars, each level repeats the one below it ten
    // times: 10 -> 100 -> 1,000 -> 10,000 chars. Small enough to be safe to actually expand in a
    // test process, large enough to prove the multiplier works if DTDs are not blocked.
    private const string BillionLaughsDoctype = """
        <!DOCTYPE root [
          <!ENTITY a "1234567890">
          <!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">
          <!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">
          <!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">
        ]>
        """;

    private const string ExternalEntityDoctype = """
        <!DOCTYPE root [ <!ENTITY xxe SYSTEM "file:///etc/passwd"> ]>
        """;

    [Fact]
    public void Billion_laughs_in_the_worksheet_is_rejected()
    {
        var hostile = XlsxBuilder.Build(sheetXml: $"""
            {BillionLaughsDoctype}
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
              <row r="1"><c r="A1" t="inlineStr"><is><t>&d;</t></is></c></row>
            </sheetData></worksheet>
            """);

        Assert.Throws<XmlException>(() => XlsxReader.Read(hostile));
    }

    [Fact]
    public void Billion_laughs_in_shared_strings_is_rejected()
    {
        var hostile = XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: $"""
                {BillionLaughsDoctype}
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t>&d;</t></si>
                </sst>
                """);

        Assert.Throws<XmlException>(() => XlsxReader.Read(hostile));
    }

    [Fact]
    public void External_entity_in_the_worksheet_is_rejected_and_its_content_never_appears()
    {
        var hostile = XlsxBuilder.Build(sheetXml: $"""
            {ExternalEntityDoctype}
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
              <row r="1"><c r="A1" t="inlineStr"><is><t>&xxe;</t></is></c></row>
            </sheetData></worksheet>
            """);

        Assert.Throws<XmlException>(() => XlsxReader.Read(hostile));
    }

    [Fact]
    public void External_entity_in_shared_strings_is_rejected_and_its_content_never_appears()
    {
        var hostile = XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: $"""
                {ExternalEntityDoctype}
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t>&xxe;</t></si>
                </sst>
                """);

        Assert.Throws<XmlException>(() => XlsxReader.Read(hostile));
    }

    // The entity reference goes in element content (an <extra> element styles.xml has no use
    // for), not an attribute value: referencing an external entity from an attribute value is
    // its own well-formedness violation regardless of DTD processing, and would throw for the
    // wrong reason.
    [Fact]
    public void External_entity_in_styles_is_rejected()
    {
        var hostile = XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1"><v>1</v></c>"""),
            stylesXml: $"""
                {ExternalEntityDoctype}
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="1"><xf numFmtId="0"/></cellXfs>
                  <extra>&xxe;</extra>
                </styleSheet>
                """);

        Assert.Throws<XmlException>(() => XlsxReader.Read(hostile));
    }

    // xl/workbook.xml is read by ReadDate1904 before ResolveFirstSheetPath runs, and that call has
    // no try/catch, so a DOCTYPE there throws XmlException straight out of Read(), same as the
    // worksheet, shared strings and styles parts.
    [Fact]
    public void External_entity_in_workbook_xml_is_rejected()
    {
        var hostile = XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1"><v>1</v></c>"""),
            workbookXml: $"""
                {ExternalEntityDoctype}
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
                  <extra>&xxe;</extra>
                </workbook>
                """);

        Assert.Throws<XmlException>(() => XlsxReader.Read(hostile));
    }

    // workbook.xml.rels is read inside ResolveFirstSheetPath's try/catch, which used to swallow
    // every failure, including a rejected DOCTYPE, and silently fall back to picking a worksheet
    // by convention instead. The catch now lets XmlException through, so this throws like the
    // other parts rather than quietly reading a different sheet than the caller asked for.
    [Fact]
    public void External_entity_in_workbook_rels_is_rejected()
    {
        var hostile = XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>safe</t></is></c>"""),
            workbookRelsXml: $"""
                {ExternalEntityDoctype}
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <extra>&xxe;</extra>
                </Relationships>
                """);

        Assert.Throws<XmlException>(() => XlsxReader.Read(hostile));
    }

    [Fact]
    public void A_normal_workbook_with_no_doctype_still_reads()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>Hello</t></is></c>""")));

        Assert.Equal("Hello", sheet.At(0, 0).Text);
    }

    // Whitespace handling must match what plain, unguarded XML loading did before this fix, since
    // that is a data question, not a security one. A <t> with only whitespace and no
    // xml:space="preserve" reads as "" both before and after; one with xml:space="preserve" keeps
    // the whitespace both before and after. Covered for shared strings and inline strings.
    [Fact]
    public void Shared_string_of_only_spaces_with_no_xml_space_reads_as_empty()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: """
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t>   </t></si>
                </sst>
                """));

        // CellValue.OfText maps an empty string to CellKind.Empty (Text is null), which is the
        // library's existing rule for any empty text, not something this fix changes.
        Assert.Equal(CellKind.Empty, sheet.At(0, 0).Kind);
    }

    [Fact]
    public void Shared_string_of_only_spaces_with_xml_space_preserve_keeps_the_spaces()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: """
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t xml:space="preserve">   </t></si>
                </sst>
                """));

        Assert.Equal("   ", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Inline_string_of_only_spaces_with_no_xml_space_reads_as_empty()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>   </t></is></c>""")));

        Assert.Equal(CellKind.Empty, sheet.At(0, 0).Kind);
    }

    [Fact]
    public void Inline_string_of_only_spaces_with_xml_space_preserve_keeps_the_spaces()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t xml:space="preserve">   </t></is></c>""")));

        Assert.Equal("   ", sheet.At(0, 0).Text);
    }

    // Mechanical guard for the rule this issue asks for: every XML part is loaded through
    // LoadPart, and nothing else in src/ is allowed to call an XML loader directly. Scans the
    // actual source files on disk, not the compiled assembly, so it catches a new call site
    // wherever it is added.
    [Fact]
    public void No_source_file_loads_XML_outside_LoadPart()
    {
        var forbidden = new[]
        {
            "XDocument.Load(", "XDocument.Parse(", "XElement.Load(", "XElement.Parse(",
            "XmlReader.Create(", "XmlDocument", "XPathDocument",
        };

        var offenders = new List<string>();
        foreach (var file in Directory.EnumerateFiles(SrcDir(), "*.cs", SearchOption.AllDirectories))
        {
            var text = File.ReadAllText(file);
            (int start, int end) loadPartRange = Path.GetFileName(file) == "XlsxReader.cs" ? LoadPartBodyRange(text) : (-1, -1);

            foreach (var pattern in forbidden)
            {
                for (var index = text.IndexOf(pattern, StringComparison.Ordinal);
                     index >= 0;
                     index = text.IndexOf(pattern, index + pattern.Length, StringComparison.Ordinal))
                {
                    var insideLoadPart = index >= loadPartRange.start && index < loadPartRange.end;
                    if (!insideLoadPart)
                        offenders.Add($"{Path.GetFileName(file)}: found '{pattern}' outside LoadPart");
                }
            }
        }

        Assert.True(offenders.Count == 0, string.Join('\n', offenders));
    }

    private static string SrcDir([CallerFilePath] string here = "")
    {
        // here is .../tests/Talaan.Tests/XlsxSecurityTests.cs; repo root is two levels up.
        var repoRoot = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(here)!, "..", ".."));
        return Path.Combine(repoRoot, "src");
    }

    /// <summary>Character offset range of the LoadPart method body, so its own XML-loading calls
    /// don't count as offenders.</summary>
    private static (int start, int end) LoadPartBodyRange(string text)
    {
        var sigIndex = text.IndexOf("XDocument LoadPart(Stream stream)", StringComparison.Ordinal);
        if (sigIndex < 0) return (-1, -1);

        var braceStart = text.IndexOf('{', sigIndex);
        var depth = 0;
        for (var i = braceStart; i < text.Length; i++)
        {
            if (text[i] == '{') depth++;
            else if (text[i] == '}' && --depth == 0) return (braceStart, i);
        }
        return (braceStart, text.Length);
    }
}
