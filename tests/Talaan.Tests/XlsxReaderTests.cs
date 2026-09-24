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
    public void Omitted_row_is_padded_so_later_rows_keep_their_position()
    {
        // row 2 is missing from sheetData (Excel omits empty rows), so row 3's content must
        // land at index 2, not shift up to index 1.
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            XlsxBuilder.Sheet(
                XlsxBuilder.RowAt(1, """<c r="A1" t="inlineStr"><is><t>top</t></is></c>"""),
                XlsxBuilder.RowAt(3, """<c r="A3" t="inlineStr"><is><t>bottom</t></is></c>"""))));

        Assert.Equal(3, sheet.RowCount);
        Assert.Equal("top", sheet.At(0, 0).Text);
        Assert.True(sheet.At(1, 0).IsBlank);
        Assert.Equal("bottom", sheet.At(2, 0).Text);
    }

    [Fact]
    public void Leading_gap_pads_rows_before_the_first_row_present()
    {
        // The sheet starts at row 3, so rows 0 and 1 must be padded blank before it.
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            XlsxBuilder.Sheet(
                XlsxBuilder.RowAt(3, """<c r="A3" t="inlineStr"><is><t>first</t></is></c>"""))));

        Assert.Equal(3, sheet.RowCount);
        Assert.True(sheet.At(0, 0).IsBlank);
        Assert.True(sheet.At(1, 0).IsBlank);
        Assert.Equal("first", sheet.At(2, 0).Text);
    }

    [Fact]
    public void Cells_still_align_within_a_column_gap_when_rows_are_padded()
    {
        // Positive control: column gaps (A1 then C1) still pad correctly once row padding exists.
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            XlsxBuilder.Sheet(
                XlsxBuilder.RowAt(1, """<c r="A1"><v>1</v></c><c r="C1"><v>3</v></c>"""),
                XlsxBuilder.RowAt(4, """<c r="A4"><v>4</v></c>"""))));

        Assert.Equal(4, sheet.RowCount);
        Assert.Equal(3, sheet.ColumnCount);
        Assert.Equal(1, sheet.At(0, 0).Number);
        Assert.True(sheet.At(0, 1).IsBlank);
        Assert.Equal(3, sheet.At(0, 2).Number);
        Assert.Equal(4, sheet.At(3, 0).Number);
    }

    [Fact]
    public void Sheet_with_no_row_gaps_is_unchanged()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            XlsxBuilder.Sheet(
                XlsxBuilder.RowAt(1, """<c r="A1"><v>1</v></c>"""),
                XlsxBuilder.RowAt(2, """<c r="A2"><v>2</v></c>"""),
                XlsxBuilder.RowAt(3, """<c r="A3"><v>3</v></c>"""))));

        Assert.Equal(3, sheet.RowCount);
        Assert.Equal(1, sheet.At(0, 0).Number);
        Assert.Equal(2, sheet.At(1, 0).Number);
        Assert.Equal(3, sheet.At(2, 0).Number);
    }

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
