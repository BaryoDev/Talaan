using System.Globalization;
using Talaan.Xlsx;
using Xunit;

namespace Talaan.Tests;

public class XlsxLimitAndSerialTests
{
    private static MemoryStream WideSparse(int rows)
    {
        var rowXml = Enumerable.Range(1, rows).Select(r => XlsxBuilder.RowAt(r, $"""<c r="XFD{r}"><v>1</v></c>"""));
        return XlsxBuilder.Build(XlsxBuilder.Sheet(rowXml.ToArray()));
    }

    [Fact]
    public void Wide_sparse_sheet_over_MaxCells_is_rejected_before_padding()
    {
        var ex = Assert.Throws<InvalidDataException>(
            () => XlsxReader.Read(WideSparse(500), new XlsxReadOptions { MaxCells = 100_000 }));
        Assert.Contains("MaxCells", ex.Message);
    }

    [Fact]
    public void Sheet_within_MaxCells_reads_normally()
    {
        var sheet = XlsxReader.Read(
            XlsxBuilder.Build(XlsxBuilder.Row("""<c r="A1"><v>1</v></c><c r="C1"><v>3</v></c>""")),
            new XlsxReadOptions { MaxCells = 3 });
        Assert.Equal(3, sheet.ColumnCount);
        Assert.Equal(3, sheet.At(0, 2).Number);
    }

    [Fact]
    public void Default_has_no_limit()
    {
        var sheet = XlsxReader.Read(WideSparse(3));
        Assert.Equal(16384, sheet.ColumnCount);
        Assert.Equal(3, sheet.RowCount);
    }

    private static CellValue ReadSerial(double serial)
    {
        var s = serial.ToString(CultureInfo.InvariantCulture);
        return XlsxReader.Read(XlsxBuilder.Build(
            sheetXml: XlsxBuilder.Row($"""<c r="A1" s="0"><v>{s}</v></c>"""),
            stylesXml: """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="1"><xf numFmtId="14"/></cellXfs>
                </styleSheet>
                """)).At(0, 0);
    }

    [Theory]
    [InlineData(1, 1900, 1, 1)]
    [InlineData(31, 1900, 1, 31)]
    [InlineData(32, 1900, 2, 1)]
    [InlineData(59, 1900, 2, 28)]
    [InlineData(61, 1900, 3, 1)]
    [InlineData(45000, 2023, 3, 15)]
    public void Serial_reads_as_Excel_shows_it(double serial, int y, int m, int d)
    {
        var cell = ReadSerial(serial);
        Assert.Equal(CellKind.Date, cell.Kind);
        Assert.Equal(new DateTime(y, m, d), cell.Date);
    }

    [Fact]
    public void Serial_60_keeps_its_previous_value_and_does_not_throw()
    {
        var cell = ReadSerial(60);
        Assert.Equal(CellKind.Date, cell.Kind);
        Assert.Equal(new DateTime(1900, 2, 28), cell.Date);
    }
}
