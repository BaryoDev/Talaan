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

    [Theory]
    [InlineData(0)]
    [InlineData(-5)]
    public void Zero_or_negative_MaxCells_rejects_the_first_cell(long max)
        => Assert.Throws<InvalidDataException>(() => XlsxReader.Read(
            XlsxBuilder.Build(XlsxBuilder.Row("""<c r="A1"><v>1</v></c>""")),
            new XlsxReadOptions { MaxCells = max }));

    [Fact]
    public void Empty_sheet_reads_with_MaxCells_zero()
    {
        var sheet = XlsxReader.Read(XlsxBuilder.Build(XlsxBuilder.Sheet()), new XlsxReadOptions { MaxCells = 0 });
        Assert.Equal(0, sheet.RowCount);
    }

    [Fact]
    public void Facade_applies_options_to_xlsx_and_ignores_them_for_csv()
    {
        Assert.Throws<InvalidDataException>(() => Spreadsheet.Read(
            WideSparse(500), SpreadsheetFormat.Xlsx, null, new XlsxReadOptions { MaxCells = 100_000 }));

        var csv = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("a,b\n1,2\n"));
        var sheet = Spreadsheet.Read(csv, SpreadsheetFormat.Csv, null, new XlsxReadOptions { MaxCells = 1 });
        Assert.Equal("2", sheet.At(1, 1).AsString());
    }

    [Fact]
    public void Existing_three_argument_calls_still_bind()
    {
        var csv = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("a\n"));
        Assert.Equal("a", Spreadsheet.Read(csv, SpreadsheetFormat.Csv, (System.Text.Encoding)null!).At(0, 0).AsString());
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

    // Pins current behaviour (documented in the README): serials 1 to 59 read one day early.
    [Theory]
    [InlineData(1, 1899, 12, 31)]
    [InlineData(59, 1900, 2, 27)]
    [InlineData(60, 1900, 2, 28)]
    [InlineData(61, 1900, 3, 1)]
    public void Serial_reads_as_before(double serial, int y, int m, int d)
    {
        var cell = ReadSerial(serial);
        Assert.Equal(CellKind.Date, cell.Kind);
        Assert.Equal(new DateTime(y, m, d), cell.Date);
    }

    private const long OneMb = 1024 * 1024;

    private static string BigText(int chars) => new('x', chars);

    private static string InlineSheet(string text) =>
        XlsxBuilder.Row($"""<c r="A1" t="inlineStr"><is><t>{text}</t></is></c>""");

    private static string SharedSheet => XlsxBuilder.Row("""<c r="A1" t="s"><v>0</v></c>""");

    private static string SharedStrings(string text) =>
        $"""<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>{text}</t></si></sst>""";

    [Fact]
    public void Sheet_part_over_MaxUncompressedBytes_is_rejected()
    {
        var ex = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(
            XlsxBuilder.Build(InlineSheet(BigText(20 * 1024 * 1024))),
            new XlsxReadOptions { MaxUncompressedBytes = OneMb }));
        Assert.Contains("MaxUncompressedBytes", ex.Message);
        Assert.Contains("xl/worksheets/sheet1.xml", ex.Message);
    }

    [Fact]
    public void Shared_strings_over_MaxUncompressedBytes_is_rejected()
    {
        var ex = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(
            XlsxBuilder.Build(SharedSheet, sharedStringsXml: SharedStrings(BigText(20 * 1024 * 1024))),
            new XlsxReadOptions { MaxUncompressedBytes = OneMb }));
        Assert.Contains("MaxUncompressedBytes", ex.Message);
        Assert.Contains("xl/sharedStrings.xml", ex.Message);
    }

    [Fact]
    public void Parts_count_against_one_budget()
    {
        // Each part is about 600 KB, so each fits under 1 MB but the two together do not.
        var text = BigText(600 * 1024);
        var stream = XlsxBuilder.Build(SharedSheet, sharedStringsXml: SharedStrings(text));
        Assert.Equal(text, XlsxReader.Read(stream, new XlsxReadOptions { MaxUncompressedBytes = 2 * OneMb }).At(0, 0).Text);

        stream = XlsxBuilder.Build(InlineSheet(text), sharedStringsXml: SharedStrings(text));
        var ex = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(
            stream, new XlsxReadOptions { MaxUncompressedBytes = OneMb }));
        Assert.Contains("MaxUncompressedBytes", ex.Message);
    }

    [Fact]
    public void Understated_stored_entry_is_caught_by_the_byte_count()
    {
        var stream = XlsxBuilder.Build(
            InlineSheet(BigText(2 * 1024 * 1024)), compression: System.IO.Compression.CompressionLevel.NoCompression);
        var bytes = stream.ToArray();
        PatchUncompressedSize(bytes, "xl/worksheets/sheet1.xml", 1000);

        var ex = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(
            new MemoryStream(bytes), new XlsxReadOptions { MaxUncompressedBytes = OneMb }));
        Assert.Contains("MaxUncompressedBytes", ex.Message);
        Assert.Contains("xl/worksheets/sheet1.xml", ex.Message);
    }

    // Lowers the uncompressed size in both the central directory and the local header of one entry.
    private static void PatchUncompressedSize(byte[] zip, string name, uint size)
    {
        var eocd = zip.AsSpan().LastIndexOf(new byte[] { 0x50, 0x4B, 0x05, 0x06 });
        var pos = (int)BitConverter.ToUInt32(zip, eocd + 16);
        while (BitConverter.ToUInt32(zip, pos) == 0x02014B50)
        {
            var nameLen = BitConverter.ToUInt16(zip, pos + 28);
            var extraLen = BitConverter.ToUInt16(zip, pos + 30);
            var commentLen = BitConverter.ToUInt16(zip, pos + 32);
            var entryName = System.Text.Encoding.UTF8.GetString(zip, pos + 46, nameLen);
            if (entryName == name)
            {
                var local = (int)BitConverter.ToUInt32(zip, pos + 42);
                BitConverter.GetBytes(size).CopyTo(zip, pos + 24);
                BitConverter.GetBytes(size).CopyTo(zip, local + 22);
                return;
            }
            pos += 46 + nameLen + extraLen + commentLen;
        }
        throw new InvalidOperationException($"No entry named {name}.");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Zero_or_negative_MaxUncompressedBytes_rejects_the_first_part(long max)
    {
        var ex = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(
            XlsxBuilder.Build(XlsxBuilder.Row("""<c r="A1"><v>1</v></c>""")),
            new XlsxReadOptions { MaxUncompressedBytes = max }));
        Assert.Contains("MaxUncompressedBytes", ex.Message);
    }

    [Fact]
    public void File_within_MaxUncompressedBytes_reads_normally()
    {
        var sheet = XlsxReader.Read(
            XlsxBuilder.Build(InlineSheet("hello")), new XlsxReadOptions { MaxUncompressedBytes = OneMb });
        Assert.Equal("hello", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Default_has_no_byte_limit()
    {
        var text = BigText(20 * 1024 * 1024);
        Assert.Equal(text, XlsxReader.Read(XlsxBuilder.Build(InlineSheet(text))).At(0, 0).Text);
    }

    [Fact]
    public void Facade_applies_MaxUncompressedBytes()
    {
        var ex = Assert.Throws<InvalidDataException>(() => Spreadsheet.Read(
            XlsxBuilder.Build(InlineSheet(BigText(2 * 1024 * 1024))), SpreadsheetFormat.Xlsx, null,
            new XlsxReadOptions { MaxUncompressedBytes = OneMb }));
        Assert.Contains("MaxUncompressedBytes", ex.Message);
    }

    [Fact]
    public void MaxCells_and_MaxUncompressedBytes_both_apply()
    {
        var options = new XlsxReadOptions { MaxCells = 100_000, MaxUncompressedBytes = OneMb };

        var cells = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(WideSparse(500), options));
        Assert.Contains("MaxCells", cells.Message);

        var bytes = Assert.Throws<InvalidDataException>(() => XlsxReader.Read(
            XlsxBuilder.Build(InlineSheet(BigText(2 * 1024 * 1024))), options));
        Assert.Contains("MaxUncompressedBytes", bytes.Message);
    }
}
