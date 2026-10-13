using System.Text;
using Talaan;
using Talaan.Csv;
using Xunit;

namespace Talaan.Tests;

public class CsvWriterTests
{
    private static SheetData Grid(params CellValue[][] rows) => new(rows.Select(r => (IReadOnlyList<CellValue>)r).ToList());
    private static CellValue T(string s) => CellValue.OfText(s);

    private static string Csv(SheetData sheet, char delimiter = ',')
    {
        var ms = new MemoryStream();
        CsvWriter.Write(ms, sheet, delimiter);
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    [Fact]
    public void Plain_fields_are_not_quoted()
    {
        Assert.Equal("a,b\r\nc,d\r\n", Csv(Grid(new[] { T("a"), T("b") }, new[] { T("c"), T("d") })));
    }

    [Theory]
    [InlineData("a,b", "\"a,b\"")]
    [InlineData("say \"hi\"", "\"say \"\"hi\"\"\"")]
    [InlineData("l1\nl2", "\"l1\nl2\"")]
    [InlineData("l1\r\nl2", "\"l1\r\nl2\"")]
    [InlineData("l1\rl2", "\"l1\rl2\"")]
    [InlineData("\"", "\"\"\"\"")]
    [InlineData(" spaced ", " spaced ")]
    public void Comma_quote_and_line_breaks_are_quoted(string value, string expected)
    {
        Assert.Equal(expected + "\r\n", Csv(Grid(new[] { T(value) })));
    }

    [Fact]
    public void Tab_delimiter_quotes_tabs_but_not_commas()
    {
        Assert.Equal("a,b\t\"c\td\"\r\n", Csv(Grid(new[] { T("a,b"), T("c\td") }), '\t'));
    }

    [Fact]
    public void Typed_cells_use_AsString_and_empty_cells_are_empty_fields()
    {
        var sheet = Grid(new[]
        {
            CellValue.OfNumber(1.5), CellValue.Empty, CellValue.OfBoolean(true),
            CellValue.OfDate(new DateTime(2024, 1, 2)), CellValue.OfDate(new DateTime(2024, 1, 2, 3, 4, 5))
        });
        Assert.Equal("1.5,,TRUE,2024-01-02,2024-01-02T03:04:05\r\n", Csv(sheet));
    }

    [Fact]
    public void Output_is_utf8_without_a_bom_and_the_stream_stays_open()
    {
        var ms = new MemoryStream();
        CsvWriter.Write(ms, Grid(new[] { T("é") }));
        Assert.True(ms.CanWrite);
        Assert.Equal(new byte[] { 0xC3, 0xA9, 0x0D, 0x0A }, ms.ToArray());
    }

    [Fact]
    public void Bad_arguments_are_rejected()
    {
        Assert.Throws<ArgumentNullException>(() => CsvWriter.Write(null!, Grid()));
        Assert.Throws<ArgumentNullException>(() => CsvWriter.Write(new MemoryStream(), null!));
        foreach (var d in new[] { '"', '\r', '\n' })
            Assert.Throws<ArgumentException>(() => CsvWriter.Write(new MemoryStream(), Grid(), d));
    }

    [Fact]
    public void Csv_round_trips_through_the_reader()
    {
        var values = new[] { "plain", "a,b", "say \"hi\"", "l1\nl2", "l1\r\nl2", "  pad  ", "00123", "<&>" };
        var ms = new MemoryStream();
        CsvWriter.Write(ms, Grid(values.Select(T).ToArray(), new[] { T("x"), CellValue.Empty, T("z") }));
        ms.Position = 0;
        var back = CsvReader.Read(ms);
        for (var i = 0; i < values.Length; i++) Assert.Equal(values[i], back.At(0, i).Text);
        Assert.Equal("x", back.At(1, 0).Text);
        Assert.Equal(CellKind.Empty, back.At(1, 1).Kind);
        Assert.Equal("z", back.At(1, 2).Text);
    }

    [Fact]
    public void WriteFile_csv_writes_only_the_first_sheet_and_round_trips()
    {
        var path = XlsxWriterTests.TempPath(".csv");
        try
        {
            Spreadsheet.WriteFile(path, new[]
            {
                new SheetData(new IReadOnlyList<CellValue>[] { new[] { T("a,b"), T("c") } }, "One"),
                new SheetData(new IReadOnlyList<CellValue>[] { new[] { T("second") } }, "Two"),
            });
            Assert.Equal("\"a,b\",c\r\n", File.ReadAllText(path));
            var back = Spreadsheet.ReadFile(path);
            Assert.Equal(1, back.RowCount);
            Assert.Equal("a,b", back.At(0, 0).Text);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void WriteFile_tsv_is_tab_delimited_first_sheet_only()
    {
        var path = XlsxWriterTests.TempPath(".tsv");
        try
        {
            Spreadsheet.WriteFile(path, new[]
            {
                Grid(new[] { T("a,b"), T("c\td"), T("e") }),
                Grid(new[] { T("second") }),
            });
            Assert.Equal("a,b\t\"c\td\"\te\r\n", File.ReadAllText(path));
            var back = Spreadsheet.ReadFile(path);
            Assert.Equal("c\td", back.At(0, 1).Text);
            Assert.Equal("e", back.At(0, 2).Text);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void WriteFile_csv_overwrites_an_existing_file()
    {
        var path = XlsxWriterTests.TempPath(".csv");
        try
        {
            File.WriteAllText(path, new string('x', 500));
            Spreadsheet.WriteFile(path, new[] { Grid(new[] { T("a") }) });
            Assert.Equal("a\r\n", File.ReadAllText(path));
        }
        finally { File.Delete(path); }
    }
}
