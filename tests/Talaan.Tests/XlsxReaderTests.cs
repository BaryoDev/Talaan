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
    public void Shared_string_with_rPh_reads_only_the_base_text()
    {
        // <rPh> holds a Japanese furigana reading alongside the base <t>. It is not part of the
        // string's text and must not be concatenated into it (issue #4).
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: """
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="1"/></si>
                </sst>
                """));

        Assert.Equal("漢字", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Shared_string_rich_text_runs_still_concatenate()
    {
        // Positive control: rich text split across several <r><t> runs is a real, supported case
        // and must still be rejoined, rPh or not.
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: """
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><r><t>Hello</t></r><r><t> World</t></r></si>
                </sst>
                """));

        Assert.Equal("Hello World", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Shared_string_rich_text_run_with_rPh_reads_only_the_base_text()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: """
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><r><t>東京</t></r><rPh sb="0" eb="2"><t>とうきょう</t></rPh></si>
                </sst>
                """));

        Assert.Equal("東京", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Shared_string_plain_t_is_unaffected_by_the_rPh_fix()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: """
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t>Plain</t></si>
                </sst>
                """));

        Assert.Equal("Plain", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Shared_string_preserves_xml_space_preserve_whitespace()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>"""),
            sharedStringsXml: """
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t xml:space="preserve">  spaced  </t></si>
                </sst>
                """));

        Assert.Equal("  spaced  ", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Inline_string_with_rPh_reads_only_the_base_text()
    {
        // Inline strings (<is>) share the same CT_Rst content model as shared strings, so they can
        // carry <rPh> too.
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row(
                """<c r="A1" t="inlineStr"><is><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh></is></c>""")));

        Assert.Equal("漢字", sheet.At(0, 0).Text);
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
    public void Row_reference_above_the_maximum_row_throws()
    {
        var stream = XlsxBuilder.Build(
            XlsxBuilder.Sheet(XlsxBuilder.RowAt(1048577, """<c r="A1048577"><v>1</v></c>""")));

        var ex = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(stream));
        Assert.Contains("1048577", ex.Message);
    }

    [Fact]
    public void Row_reference_at_the_maximum_row_still_reads()
    {
        // Boundary case for the cap above: the last legal row still works, padding the full grid.
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            XlsxBuilder.Sheet(XlsxBuilder.RowAt(1048576, """<c r="A1048576"><v>1</v></c>"""))));

        Assert.Equal(1_048_576, sheet.RowCount);
        Assert.Equal(1, sheet.At(1_048_575, 0).Number);
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
