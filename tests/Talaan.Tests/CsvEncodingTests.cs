using System.Text;
using Talaan;
using Talaan.Csv;
using Xunit;

namespace Talaan.Tests;

public class CsvEncodingTests
{
    // Windows-1252 / Latin-1 bytes for "naivé": 6E 61 69 76 E9.
    private static readonly byte[] Latin1NaiveBytes = { 0x6E, 0x61, 0x69, 0x76, 0xE9 };

    [Fact]
    public void CsvReader_reads_a_single_byte_encoding_when_told_which_one()
    {
        var sheet = CsvReader.Read(new MemoryStream(Latin1NaiveBytes), Encoding.Latin1);
        Assert.Equal("naivé", sheet.At(0, 0).Text);
    }

    [Fact]
    public void CsvReader_without_an_encoding_still_decodes_as_utf8_by_default()
    {
        var sheet = CsvReader.Read(new MemoryStream(Latin1NaiveBytes));
        Assert.Equal("naiv�", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Spreadsheet_Read_with_format_and_encoding_reads_a_single_byte_encoding()
    {
        var sheet = Spreadsheet.Read(new MemoryStream(Latin1NaiveBytes), SpreadsheetFormat.Csv, Encoding.Latin1);
        Assert.Equal("naivé", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Spreadsheet_ReadFile_with_encoding_reads_a_single_byte_encoding()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".csv");
        try
        {
            File.WriteAllBytes(path, Latin1NaiveBytes);
            var sheet = Spreadsheet.ReadFile(path, Encoding.Latin1);
            Assert.Equal("naivé", sheet.At(0, 0).Text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_utf8_bom_file_still_reads_with_the_default()
    {
        var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes("naivé,x\n")).ToArray();

        var sheet = CsvReader.Read(new MemoryStream(bytes));

        Assert.Equal("naivé", sheet.At(0, 0).Text);
    }

    // A byte-order mark is an explicit, unambiguous signal, so it overrides an encoding the caller
    // passed. detectEncodingFromByteOrderMarks stays true on the new overload, matching the
    // no-encoding one: a UTF-16 file decodes correctly even if the caller guessed Latin1.
    [Fact]
    public void A_utf16_bom_wins_over_an_explicit_encoding()
    {
        var bytes = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("naivé,x\n")).ToArray();

        var sheet = CsvReader.Read(new MemoryStream(bytes), Encoding.Latin1);

        Assert.Equal("naivé", sheet.At(0, 0).Text);
    }

    [Fact]
    public void CsvReader_encoding_overload_leaves_the_callers_stream_open()
    {
        var stream = new MemoryStream(Latin1NaiveBytes);

        CsvReader.Read(stream, Encoding.Latin1);

        Assert.True(stream.CanRead);
        stream.Seek(0, SeekOrigin.Begin);
        var sheet = CsvReader.Read(stream, Encoding.Latin1);
        Assert.Equal("naivé", sheet.At(0, 0).Text);
    }

    [Fact]
    public void Spreadsheet_Read_encoding_overload_leaves_the_callers_stream_open()
    {
        var stream = new MemoryStream(Latin1NaiveBytes);

        Spreadsheet.Read(stream, SpreadsheetFormat.Csv, Encoding.Latin1);

        Assert.True(stream.CanRead);
        stream.Seek(0, SeekOrigin.Begin);
        var sheet = Spreadsheet.Read(stream, SpreadsheetFormat.Csv, Encoding.Latin1);
        Assert.Equal("naivé", sheet.At(0, 0).Text);
    }
}
