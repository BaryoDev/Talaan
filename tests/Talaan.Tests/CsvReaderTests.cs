using System.Text;
using Talaan;
using Talaan.Csv;
using Xunit;

namespace Talaan.Tests;

public class CsvReaderTests
{
    private static SheetData Parse(string csv)
        => CsvReader.Read(new MemoryStream(Encoding.UTF8.GetBytes(csv)));

    [Fact]
    public void Reads_simple_rows()
    {
        var sheet = Parse("a,b,c\n1,2,3\n");
        Assert.Equal(2, sheet.RowCount);
        Assert.Equal("a", sheet.At(0, 0).AsString());
        Assert.Equal("3", sheet.At(1, 2).AsString());
    }

    [Fact]
    public void Handles_quoted_fields_with_commas_and_quotes()
    {
        var sheet = Parse("\"Doe, John\",\"say \"\"hi\"\"\"\n");
        Assert.Equal("Doe, John", sheet.At(0, 0).AsString());
        Assert.Equal("say \"hi\"", sheet.At(0, 1).AsString());
    }

    [Fact]
    public void Handles_newline_inside_quoted_field()
    {
        var sheet = Parse("\"line1\nline2\",b\n");
        Assert.Equal(1, sheet.RowCount);
        Assert.Equal("line1\nline2", sheet.At(0, 0).AsString());
        Assert.Equal("b", sheet.At(0, 1).AsString());
    }

    [Fact]
    public void Empty_fields_are_blank()
    {
        var sheet = Parse("a,,c\n");
        Assert.True(sheet.At(0, 1).IsBlank);
    }

    [Fact]
    public void Trailing_newline_does_not_add_empty_row()
    {
        var sheet = Parse("a,b\nc,d\n");
        Assert.Equal(2, sheet.RowCount);
    }

    [Fact]
    public void Unterminated_quote_is_rejected()
    {
        var ex = Assert.Throws<InvalidDataException>(() => Parse("a,\"b,c\nd,e"));
        Assert.Contains("record 1", ex.Message);
    }

    [Fact]
    public void Unterminated_quote_names_the_record_it_started_in()
    {
        var ex = Assert.Throws<InvalidDataException>(() => Parse("a,b\nc,\"d,e"));
        Assert.Contains("record 2", ex.Message);
    }

    [Fact]
    public void Quoted_field_closed_at_end_of_file_without_newline_is_fine()
        => Assert.Equal("b", Parse("a,\"b\"").At(0, 1).AsString());

    [Fact]
    public void Quoted_field_containing_the_delimiter_reads_correctly()
    {
        var sheet = Parse("\"a,b\",c\n");
        Assert.Equal("a,b", sheet.At(0, 0).AsString());
        Assert.Equal("c", sheet.At(0, 1).AsString());
    }

    [Fact]
    public void Quoted_field_containing_a_newline_reads_correctly()
    {
        var sheet = Parse("\"a\nb\",c\n");
        Assert.Equal("a\nb", sheet.At(0, 0).AsString());
        Assert.Equal("c", sheet.At(0, 1).AsString());
    }

    [Fact]
    public void Escaped_quote_inside_a_quoted_field_reads_correctly()
    {
        var sheet = Parse("\"a\"\"b\",c\n");
        Assert.Equal("a\"b", sheet.At(0, 0).AsString());
        Assert.Equal("c", sheet.At(0, 1).AsString());
    }
}
