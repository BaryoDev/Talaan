using System.Globalization;
using Talaan;
using Talaan.Xlsx;
using Xunit;

namespace Talaan.Tests;

public class XlsxReaderTests
{
    [Theory]
    [InlineData("A1", 0)]
    [InlineData("B1", 1)]
    [InlineData("Z9", 25)]
    [InlineData("AA1", 26)]
    [InlineData("AB12", 27)]
    public void Column_index_from_reference(string cellRef, int expected)
        => Assert.Equal(expected, XlsxReader.ColumnIndex(cellRef));

    [Fact]
    public void Shared_string_resolves_through_the_shared_strings_table()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: """
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t>Hello</t></si>
                </sst>
                """));

        Assert.Equal("Hello", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Inline_string_reads_its_text()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>World</t></is></c>""")));

        Assert.Equal("World", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Number_cell_reads_as_a_number()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1"><v>42.5</v></c>""")));

        Assert.Equal(CellKind.Number, sheet.At(0, 0).Kind);
        Assert.Equal(42.5, sheet.At(0, 0).Number);
    }

    [Theory]
    [InlineData("1", true)]
    [InlineData("0", false)]
    public void Boolean_cell_reads_true_and_false(string raw, bool expected)
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1" t="b"><v>{raw}</v></c>""")));

        Assert.Equal(CellKind.Boolean, sheet.At(0, 0).Kind);
        Assert.Equal(expected, sheet.At(0, 0).Boolean);
    }

    [Fact]
    public void Number_with_a_builtin_date_style_reads_as_a_date()
    {
        var date = new DateTime(2024, 3, 15);
        var serial = date.ToOADate().ToString(CultureInfo.InvariantCulture);

        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1" s="0"><v>{serial}</v></c>"""),
            stylesXml: """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="1"><xf numFmtId="14"/></cellXfs>
                </styleSheet>
                """));

        Assert.Equal(CellKind.Date, sheet.At(0, 0).Kind);
        Assert.Equal(date, sheet.At(0, 0).Date);
    }

    [Fact]
    public void Number_with_a_custom_date_format_reads_as_a_date()
    {
        var date = new DateTime(2026, 9, 1);
        var serial = date.ToOADate().ToString(CultureInfo.InvariantCulture);

        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1" s="0"><v>{serial}</v></c>"""),
            stylesXml: """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
                  <cellXfs count="1"><xf numFmtId="164"/></cellXfs>
                </styleSheet>
                """));

        Assert.Equal(CellKind.Date, sheet.At(0, 0).Kind);
        Assert.Equal(date, sheet.At(0, 0).Date);
    }

    [Fact]
    public void Date1904_serial_reads_as_the_correct_date()
    {
        // The 1904 date system counts from 1904-01-01. Its serials are 1462 days behind the 1900
        // system OADate uses, so writing the date as a 1904 serial without offsetting reads 1462
        // days early.
        var date = new DateTime(2024, 3, 15);
        var serial = (date.ToOADate() - 1462).ToString(CultureInfo.InvariantCulture);

        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1" s="0"><v>{serial}</v></c>"""),
            stylesXml: """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="1"><xf numFmtId="14"/></cellXfs>
                </styleSheet>
                """,
            workbookXml: Date1904WorkbookXml("1")));

        Assert.Equal(CellKind.Date, sheet.At(0, 0).Kind);
        Assert.Equal(date, sheet.At(0, 0).Date);
    }

    [Fact]
    public void Date1904_true_attribute_reads_as_the_correct_date()
    {
        // date1904 is xsd:boolean, so "true" is as valid as "1".
        var date = new DateTime(2024, 3, 15);
        var serial = (date.ToOADate() - 1462).ToString(CultureInfo.InvariantCulture);

        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1" s="0"><v>{serial}</v></c>"""),
            stylesXml: """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="1"><xf numFmtId="14"/></cellXfs>
                </styleSheet>
                """,
            workbookXml: Date1904WorkbookXml("true")));

        Assert.Equal(CellKind.Date, sheet.At(0, 0).Kind);
        Assert.Equal(date, sheet.At(0, 0).Date);
    }

    [Fact]
    public void Date1904_explicit_zero_keeps_1900_behaviour()
    {
        var date = new DateTime(2024, 3, 15);
        var serial = date.ToOADate().ToString(CultureInfo.InvariantCulture);

        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1" s="0"><v>{serial}</v></c>"""),
            stylesXml: """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="1"><xf numFmtId="14"/></cellXfs>
                </styleSheet>
                """,
            workbookXml: Date1904WorkbookXml("0")));

        Assert.Equal(CellKind.Date, sheet.At(0, 0).Kind);
        Assert.Equal(date, sheet.At(0, 0).Date);
    }

    [Fact]
    public void Date1904_absent_keeps_1900_behaviour()
    {
        var date = new DateTime(2024, 3, 15);
        var serial = date.ToOADate().ToString(CultureInfo.InvariantCulture);

        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1" s="0"><v>{serial}</v></c>"""),
            stylesXml: """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="1"><xf numFmtId="14"/></cellXfs>
                </styleSheet>
                """,
            workbookXml: Date1904WorkbookXml(date1904Attr: null)));

        Assert.Equal(CellKind.Date, sheet.At(0, 0).Kind);
        Assert.Equal(date, sheet.At(0, 0).Date);
    }

    [Fact]
    public void Date1904_with_a_time_component_keeps_the_time()
    {
        var date = new DateTime(2024, 3, 15, 13, 45, 0);
        var serial = (date.ToOADate() - 1462).ToString(CultureInfo.InvariantCulture);

        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1" s="0"><v>{serial}</v></c>"""),
            stylesXml: """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="1"><xf numFmtId="22"/></cellXfs>
                </styleSheet>
                """,
            workbookXml: Date1904WorkbookXml("1")));

        Assert.Equal(CellKind.Date, sheet.At(0, 0).Kind);
        Assert.Equal(date, sheet.At(0, 0).Date);
    }

    /// <summary>A default single-sheet workbook, optionally with a `workbookPr date1904` attribute.</summary>
    private static string Date1904WorkbookXml(string? date1904Attr) => $"""
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          {(date1904Attr is null ? "" : $"""<workbookPr date1904="{date1904Attr}"/>""")}
          <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
        """;

    [Fact]
    public void First_sheet_resolves_through_workbook_rels_not_alphabetical_fallback()
    {
        // Two worksheet parts exist. "aaa_decoy.xml" sorts first alphabetically, so this only
        // passes if the reader follows workbook.xml -> workbook.xml.rels -> "real.xml" rather than
        // falling back to the alphabetically-first part under xl/worksheets/.
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>real</t></is></c>"""),
            workbookXml: """
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId9"/></sheets>
                </workbook>
                """,
            workbookRelsXml: """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/real.xml"/>
                </Relationships>
                """,
            sheetPath: "xl/worksheets/real.xml",
            extraParts: new Dictionary<string, string>
            {
                ["xl/worksheets/aaa_decoy.xml"] =
                    XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>decoy</t></is></c>""")
            }));

        Assert.Equal("real", sheet.At(0, 0).Text);
    }
}
