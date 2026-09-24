using System.Xml;
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

    // ResolveFirstSheetPath wraps workbook.xml and workbook.xml.rels in a try/catch that already
    // falls back to the worksheets/ convention on any failure (a pre-existing design for a
    // missing or malformed workbook.xml). A DOCTYPE there is rejected at the same LoadPart call
    // as the other parts, but the caller sees that as a safe fallback, not a thrown XmlException:
    // the malicious part is never used and no entity content reaches a cell.
    [Fact]
    public void External_entity_in_workbook_xml_does_not_leak_and_falls_back_to_convention()
    {
        var xlsx = XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>safe</t></is></c>"""),
            workbookXml: $"""
                {ExternalEntityDoctype}
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
                  <extra>&xxe;</extra>
                </workbook>
                """);

        var sheet = XlsxReader.Read(xlsx);

        Assert.Equal("safe", sheet.At(0, 0).Text);
    }

    [Fact]
    public void External_entity_in_workbook_rels_does_not_leak_and_falls_back_to_convention()
    {
        var xlsx = XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>safe</t></is></c>"""),
            workbookRelsXml: $"""
                {ExternalEntityDoctype}
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <extra>&xxe;</extra>
                </Relationships>
                """);

        var sheet = XlsxReader.Read(xlsx);

        Assert.Equal("safe", sheet.At(0, 0).Text);
    }

    [Fact]
    public void A_normal_workbook_with_no_doctype_still_reads()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>Hello</t></is></c>""")));

        Assert.Equal("Hello", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Shared_string_of_only_spaces_is_preserved()
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
}
