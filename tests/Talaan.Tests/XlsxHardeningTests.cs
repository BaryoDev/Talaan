using Talaan;
using Talaan.Xlsx;
using Xunit;

namespace Talaan.Tests;

public class XlsxHardeningTests
{
    private const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static string NestedSheet(int depth) =>
        $"""<worksheet xmlns="{Ns}"><sheetData>""" +
        string.Concat(Enumerable.Repeat("<x>", depth)) +
        string.Concat(Enumerable.Repeat("</x>", depth)) +
        "</sheetData></worksheet>";

    [Fact]
    public void Part_nested_just_past_the_limit_is_rejected()
    {
        var hostile = XlsxBuilder.Build(sheetXml: NestedSheet(300));

        var ex = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(hostile));
        Assert.Contains("xl/worksheets/sheet1.xml", ex.Message);
    }

    [Fact]
    public void Very_deep_part_is_rejected_quickly()
    {
        var hostile = XlsxBuilder.Build(sheetXml: NestedSheet(100_000));

        var watch = System.Diagnostics.Stopwatch.StartNew();
        Assert.Throws<InvalidDataException>(() => XlsxReader.Read(hostile));
        Assert.True(watch.ElapsedMilliseconds < 2000);
    }

    [Fact]
    public void Normal_workbook_still_reads()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>hi</t></is></c><c r="B1"><v>2</v></c>""")));

        Assert.Equal("hi", sheet.At(0, 0).Text);
        Assert.Equal(2, sheet.At(0, 1).Number);
    }

    [Theory]
    [InlineData("""<numFmts count="1"><numFmt numFmtId="x" formatCode="yyyy"/></numFmts><cellXfs count="1"><xf numFmtId="0"/></cellXfs>""")]
    [InlineData("""<cellXfs count="1"><xf numFmtId="abc"/></cellXfs>""")]
    [InlineData("""<cellXfs count="1"><xf numFmtId="99999999999"/></cellXfs>""")]
    public void Non_numeric_numFmtId_is_invalid_data(string body)
    {
        var hostile = XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1"><v>1</v></c>"""),
            stylesXml: $"""<styleSheet xmlns="{Ns}">{body}</styleSheet>""");

        Assert.Throws<InvalidDataException>(() => XlsxReader.Read(hostile));
    }

    private static MemoryStream WithTarget(string target) => XlsxBuilder.Build(
        sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>found</t></is></c>"""),
        workbookRelsXml: $"""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{target}"/>
            </Relationships>
            """);

    [Theory]
    [InlineData("../xl/worksheets/sheet1.xml")]
    [InlineData("./worksheets/sheet1.xml")]
    [InlineData("/xl/worksheets/sheet1.xml")]
    [InlineData("worksheets/../worksheets/sheet1.xml")]
    public void Relationship_target_with_dot_segments_resolves(string target)
    {
        var sheet = XlsxReader.Read(WithTarget(target));

        Assert.Equal("found", sheet.At(0, 0).Text);
    }

    [Theory]
    [InlineData("../../x.xml")]
    [InlineData("../../../etc/passwd")]
    public void Relationship_target_escaping_the_package_is_rejected(string target)
    {
        var ex = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(WithTarget(target)));
        Assert.Contains("escapes the package", ex.Message);
    }

    [Fact]
    public void External_relationship_is_not_resolved_as_a_package_part()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row("""<c r="A1" t="inlineStr"><is><t>found</t></is></c>"""),
            workbookRelsXml: """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../../x.xml" TargetMode="External"/>
                </Relationships>
                """));

        Assert.Equal("found", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Part_at_the_limit_reads_and_one_deeper_is_rejected()
    {
        // worksheet and sheetData take two levels, so 254 x elements put the deepest at depth 255.
        Assert.Empty(XlsxReader.Read(XlsxBuilder.Build(sheetXml: NestedSheet(254))).Rows);
        Assert.Throws<InvalidDataException>(() => XlsxReader.Read(XlsxBuilder.Build(sheetXml: NestedSheet(255))));
    }

    [Theory]
    [InlineData("1,5")]
    [InlineData("1,000")]
    [InlineData("$5")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("-Infinity")]
    public void Non_invariant_or_non_finite_number_stays_text(string raw)
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1"><v>{raw}</v></c>""")));

        Assert.Equal(CellKind.Text, sheet.At(0, 0).Kind);
        Assert.Equal(raw, sheet.At(0, 0).Text);
    }

    [Theory]
    [InlineData("1.5", 1.5)]
    [InlineData("-3", -3)]
    [InlineData("1E3", 1000)]
    [InlineData("1.25e-2", 0.0125)]
    public void Plain_invariant_numbers_still_read(string raw, double expected)
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1"><v>{raw}</v></c>""")));

        Assert.Equal(CellKind.Number, sheet.At(0, 0).Kind);
        Assert.Equal(expected, sheet.At(0, 0).Number);
    }
}
