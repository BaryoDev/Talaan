using System.IO.Compression;
using System.Text.RegularExpressions;
using Talaan;
using Xunit;

namespace Talaan.Tests;

public class XlsxWriterTests
{
    private static SheetData Grid(params CellValue[][] rows) => new(rows.Select(r => (IReadOnlyList<CellValue>)r).ToList());

    private static SheetData Named(string? name, params CellValue[][] rows) =>
        new(rows.Select(r => (IReadOnlyList<CellValue>)r).ToList(), name);

    private static CellValue T(string s) => CellValue.OfText(s);
    private static CellValue N(double d) => CellValue.OfNumber(d);
    private static CellValue B(bool b) => CellValue.OfBoolean(b);
    private static CellValue D(DateTime d) => CellValue.OfDate(d);
    private static readonly CellValue E = CellValue.Empty;

    private static byte[] Write(params SheetData[] sheets)
    {
        var ms = new MemoryStream();
        Spreadsheet.WriteXlsx(ms, sheets);
        return ms.ToArray();
    }

    private static SheetData RoundTrip(SheetData sheet)
    {
        using var ms = new MemoryStream(Write(sheet));
        return Spreadsheet.Read(ms, SpreadsheetFormat.Xlsx);
    }

    private static string Part(byte[] xlsx, string path)
    {
        using var zip = new ZipArchive(new MemoryStream(xlsx));
        using var r = new StreamReader(zip.GetEntry(path)!.Open());
        return r.ReadToEnd();
    }

    private static List<string> SheetNames(byte[] xlsx) =>
        Regex.Matches(Part(xlsx, "xl/workbook.xml"), "<sheet name=\"([^\"]*)\"")
            .Select(m => m.Groups[1].Value).ToList();

    // The reader takes the first sheet only, so point the first relationship at sheet N and read that.
    private static SheetData ReadSheetAt(byte[] xlsx, int n)
    {
        var rels = Part(xlsx, "xl/_rels/workbook.xml.rels")
            .Replace("worksheets/sheet1.xml", $"worksheets/sheet{n}.xml");
        var ms = new MemoryStream();
        using (var src = new ZipArchive(new MemoryStream(xlsx)))
        using (var dst = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var e in src.Entries)
            {
                using var w = dst.CreateEntry(e.FullName).Open();
                if (e.FullName == "xl/_rels/workbook.xml.rels")
                    w.Write(System.Text.Encoding.UTF8.GetBytes(rels));
                else
                    e.Open().CopyTo(w);
            }
        }
        ms.Position = 0;
        return Spreadsheet.Read(ms, SpreadsheetFormat.Xlsx);
    }

    // ---- round trips ----

    [Fact]
    public void Text_round_trips_including_markup_and_spacing()
    {
        var values = new[] { "plain", "a,b", "say \"hi\"", "1 < 2 & 3 > 2", "  lead", "trail  ", "  both  ", "line1\nline2", "tab\there", "naïve 日本 😀" };
        var back = RoundTrip(Grid(values.Select(T).ToArray()));
        for (var i = 0; i < values.Length; i++)
        {
            Assert.Equal(CellKind.Text, back.At(0, i).Kind);
            Assert.Equal(values[i], back.At(0, i).Text);
        }
    }

    [Fact]
    public void Carriage_return_in_text_survives()
    {
        var back = RoundTrip(Grid(new[] { T("a\r\nb"), T("c\rd") }));
        Assert.Equal("a\r\nb", back.At(0, 0).Text);
        Assert.Equal("c\rd", back.At(0, 1).Text);
    }

    [Fact]
    public void Number_looking_text_stays_text()
    {
        var back = RoundTrip(Grid(new[] { T("00123"), T("1e5"), T("TRUE"), T("2024-01-02") }));
        for (var i = 0; i < 4; i++) Assert.Equal(CellKind.Text, back.At(0, i).Kind);
        Assert.Equal("00123", back.At(0, 0).Text);
    }

    [Fact]
    public void Numbers_round_trip_exactly()
    {
        var values = new[] { 0d, 1, -1, 42, 3.14159, -0.5, 0.1 + 0.2, 1234567890123, 1e21, -1e-9, double.MaxValue, double.Epsilon, 123456789.123456789 };
        var back = RoundTrip(Grid(values.Select(N).ToArray()));
        for (var i = 0; i < values.Length; i++)
        {
            Assert.Equal(CellKind.Number, back.At(0, i).Kind);
            Assert.Equal(values[i], back.At(0, i).Number);
        }
    }

    [Fact]
    public void Booleans_round_trip()
    {
        var back = RoundTrip(Grid(new[] { B(true), B(false) }));
        Assert.Equal(CellKind.Boolean, back.At(0, 0).Kind);
        Assert.True(back.At(0, 0).Boolean);
        Assert.False(back.At(0, 1).Boolean);
    }

    [Fact]
    public void Dates_round_trip_as_dates()
    {
        var day = new DateTime(2024, 2, 29);
        var stamp = new DateTime(2024, 12, 31, 23, 59, 58);
        var back = RoundTrip(Grid(new[] { D(day), D(stamp) }));
        Assert.Equal(CellKind.Date, back.At(0, 0).Kind);
        Assert.Equal(day, back.At(0, 0).Date);
        Assert.Equal(CellKind.Date, back.At(0, 1).Kind);
        Assert.Equal(stamp, back.At(0, 1).Date);
    }

    [Fact]
    public void Date_without_time_uses_a_date_only_format_and_with_time_a_datetime_format()
    {
        var styles = Part(Write(Grid(new[] { D(new DateTime(2024, 1, 1)) })), "xl/styles.xml");
        Assert.Contains("yyyy", styles);
        Assert.Contains("hh:mm:ss", styles);
    }

    [Fact]
    public void Date_before_1900_03_01_is_rejected()
    {
        Assert.Throws<ArgumentException>(() => Write(Grid(new[] { D(new DateTime(1900, 1, 1)) })));
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void Non_finite_numbers_are_rejected(double value)
    {
        Assert.Throws<ArgumentException>(() => Write(Grid(new[] { N(value) })));
    }

    [Fact]
    public void Empty_cells_in_the_middle_keep_columns_aligned()
    {
        var back = RoundTrip(Grid(new[] { T("a"), E, E, T("d") }));
        Assert.Equal("a", back.At(0, 0).Text);
        Assert.Equal(CellKind.Empty, back.At(0, 1).Kind);
        Assert.Equal(CellKind.Empty, back.At(0, 2).Kind);
        Assert.Equal("d", back.At(0, 3).Text);
    }

    [Fact]
    public void Empty_text_is_skipped()
    {
        var back = RoundTrip(Grid(new[] { T("a"), new CellValue(CellKind.Text, "", null, null, null), T("c") }));
        Assert.Equal(CellKind.Empty, back.At(0, 1).Kind);
        Assert.Equal("c", back.At(0, 2).Text);
    }

    [Fact]
    public void Ragged_rows_and_empty_rows_keep_their_positions()
    {
        var back = RoundTrip(Grid(
            new[] { T("a"), T("b"), T("c") },
            Array.Empty<CellValue>(),
            new[] { E, E },
            new[] { N(1) }));
        Assert.Equal(4, back.RowCount);
        Assert.Equal("c", back.At(0, 2).Text);
        Assert.Equal(CellKind.Empty, back.At(1, 0).Kind);
        Assert.Equal(CellKind.Empty, back.At(2, 0).Kind);
        Assert.Equal(1d, back.At(3, 0).Number);
    }

    [Fact]
    public void Columns_past_Z_are_addressed_correctly()
    {
        var row = new CellValue[704];
        for (var i = 0; i < row.Length; i++) row[i] = E;
        row[25] = T("Z");
        row[26] = T("AA");
        row[51] = T("AZ");
        row[52] = T("BA");
        row[701] = T("ZZ");
        row[702] = T("AAA");
        var bytes = Write(Grid(row));
        var xml = Part(bytes, "xl/worksheets/sheet1.xml");
        foreach (var r in new[] { "Z1", "AA1", "AZ1", "BA1", "ZZ1", "AAA1" })
            Assert.Contains($"r=\"{r}\"", xml);

        var back = ReadSheetAt(bytes, 1);
        Assert.Equal("Z", back.At(0, 25).Text);
        Assert.Equal("AA", back.At(0, 26).Text);
        Assert.Equal("AZ", back.At(0, 51).Text);
        Assert.Equal("BA", back.At(0, 52).Text);
        Assert.Equal("ZZ", back.At(0, 701).Text);
        Assert.Equal("AAA", back.At(0, 702).Text);
    }

    [Fact]
    public void Last_column_XFD_is_allowed_and_one_more_is_not()
    {
        var ok = new CellValue[16384];
        for (var i = 0; i < ok.Length; i++) ok[i] = E;
        ok[16383] = T("x");
        Assert.Equal("x", RoundTrip(Grid(ok)).At(0, 16383).Text);

        Assert.Throws<ArgumentException>(() => Write(Grid(new CellValue[16385])));
    }

    // ---- package shape ----

    [Fact]
    public void Package_has_the_expected_parts_and_every_part_is_well_formed()
    {
        var bytes = Write(Named("A", new[] { T("x") }), Named("B", new[] { T("y") }));
        using var zip = new ZipArchive(new MemoryStream(bytes));
        var names = zip.Entries.Select(e => e.FullName).OrderBy(n => n, StringComparer.Ordinal).ToArray();
        Assert.Equal(new[]
        {
            "[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.xml.rels", "xl/styles.xml",
            "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"
        }, names);
        foreach (var e in zip.Entries)
        {
            using var s = e.Open();
            var doc = System.Xml.Linq.XDocument.Load(s);
            Assert.NotNull(doc.Root);
        }
    }

    [Fact]
    public void Text_uses_inline_strings_with_preserved_space_and_no_shared_strings()
    {
        var bytes = Write(Grid(new[] { T(" a ") }));
        var xml = Part(bytes, "xl/worksheets/sheet1.xml");
        Assert.Contains("t=\"inlineStr\"", xml);
        Assert.Contains("xml:space=\"preserve\"", xml);
        using var zip = new ZipArchive(new MemoryStream(bytes));
        Assert.Null(zip.GetEntry("xl/sharedStrings.xml"));
    }

    [Fact]
    public void Booleans_are_written_as_t_b_with_1_or_0()
    {
        var xml = Part(Write(Grid(new[] { B(true), B(false) })), "xl/worksheets/sheet1.xml");
        Assert.Matches("<c r=\"A1\" t=\"b\"><v>1</v></c>", xml);
        Assert.Matches("<c r=\"B1\" t=\"b\"><v>0</v></c>", xml);
    }

    [Fact]
    public void Numbers_use_the_invariant_culture()
    {
        var old = System.Globalization.CultureInfo.CurrentCulture;
        try
        {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("de-DE");
            var xml = Part(Write(Grid(new[] { N(1.5) })), "xl/worksheets/sheet1.xml");
            Assert.Contains("<v>1.5</v>", xml);
        }
        finally { System.Globalization.CultureInfo.CurrentCulture = old; }
    }

    // ---- multiple sheets ----

    [Fact]
    public void Multiple_sheets_keep_their_names_order_and_content()
    {
        var bytes = Write(
            Named("People", new[] { T("name"), T("age") }, new[] { T("Ana"), N(30) }),
            Named("Flags", new[] { B(true) }),
            Named("Dates", new[] { D(new DateTime(2024, 5, 6)) }));

        Assert.Equal(new[] { "People", "Flags", "Dates" }, SheetNames(bytes));

        var second = ReadSheetAt(bytes, 2);
        Assert.True(second.At(0, 0).Boolean);
        var third = ReadSheetAt(bytes, 3);
        Assert.Equal(new DateTime(2024, 5, 6), third.At(0, 0).Date);
        var first = ReadSheetAt(bytes, 1);
        Assert.Equal("Ana", first.At(1, 0).Text);
        Assert.Equal(30d, first.At(1, 1).Number);
    }

    [Fact]
    public void Sheet_names_with_markup_are_escaped()
    {
        var bytes = Write(Named("A&B <1>", new[] { T("x") }));
        Assert.Equal(new[] { "A&amp;B &lt;1&gt;" }, SheetNames(bytes));
    }

    // ---- sheet names ----

    [Theory]
    [InlineData(null, "Sheet1")]
    [InlineData("", "Sheet1")]
    [InlineData("   ", "Sheet1")]
    public void Blank_name_gets_a_default(string? name, string expected)
    {
        Assert.Equal(new[] { expected }, SheetNames(Write(Named(name, new[] { T("x") }))));
    }

    [Fact]
    public void Default_name_uses_the_sheets_position()
    {
        var bytes = Write(Named("First", new[] { T("x") }), Named(null, new[] { T("y") }), Named("Third", new[] { T("z") }), Named("", new[] { T("w") }));
        Assert.Equal(new[] { "First", "Sheet2", "Third", "Sheet4" }, SheetNames(bytes));
    }

    [Fact]
    public void Name_of_31_characters_is_allowed_and_32_is_not()
    {
        Assert.Equal(new[] { new string('a', 31) }, SheetNames(Write(Named(new string('a', 31), new[] { T("x") }))));
        var ex = Assert.Throws<ArgumentException>(() => Write(Named(new string('a', 32), new[] { T("x") })));
        Assert.Contains("31", ex.Message);
    }

    [Theory]
    [InlineData('[')]
    [InlineData(']')]
    [InlineData(':')]
    [InlineData('*')]
    [InlineData('?')]
    [InlineData('/')]
    [InlineData('\\')]
    public void Each_forbidden_character_is_rejected(char c)
    {
        var ex = Assert.Throws<ArgumentException>(() => Write(Named($"a{c}b", new[] { T("x") })));
        Assert.Contains(c.ToString(), ex.Message);
    }

    [Fact]
    public void Leading_or_trailing_apostrophe_is_rejected()
    {
        Assert.Throws<ArgumentException>(() => Write(Named("'a", new[] { T("x") })));
        Assert.Throws<ArgumentException>(() => Write(Named("a'", new[] { T("x") })));
        Assert.Equal(new[] { "it's" }, SheetNames(Write(Named("it's", new[] { T("x") }))));
    }

    [Theory]
    [InlineData("Data", "Data")]
    [InlineData("Data", "data")]
    [InlineData("Data", "DATA")]
    public void Duplicate_names_are_rejected_case_insensitively(string a, string b)
    {
        var ex = Assert.Throws<ArgumentException>(() => Write(Named(a, new[] { T("x") }), Named(b, new[] { T("y") })));
        Assert.Contains("Duplicate", ex.Message);
    }

    [Fact]
    public void Explicit_name_that_matches_a_default_is_a_duplicate()
    {
        Assert.Throws<ArgumentException>(() => Write(Named(null, new[] { T("x") }), Named("sheet1", new[] { T("y") })));
    }

    [Fact]
    public void Empty_list_is_rejected()
    {
        Assert.Throws<ArgumentException>(() => Spreadsheet.WriteXlsx(new MemoryStream(), Array.Empty<SheetData>()));
    }

    [Fact]
    public void Nulls_are_rejected()
    {
        Assert.Throws<ArgumentNullException>(() => Spreadsheet.WriteXlsx(null!, new[] { Grid() }));
        Assert.Throws<ArgumentNullException>(() => Spreadsheet.WriteXlsx(new MemoryStream(), null!));
        Assert.Throws<ArgumentException>(() => Spreadsheet.WriteXlsx(new MemoryStream(), new SheetData[] { null! }));
    }

    // ---- control characters ----

    [Fact]
    public void Illegal_control_characters_are_dropped_from_text()
    {
        var dirty = "a\u0000b\u0001c\u0008d\u000Be\u000Cf\u000Eg\u001Fh￾i￿j";
        var back = RoundTrip(Grid(new[] { T(dirty) }));
        Assert.Equal("abcdefghij", back.At(0, 0).Text);
    }

    [Fact]
    public void Legal_whitespace_controls_and_surrogate_pairs_are_kept()
    {
        var back = RoundTrip(Grid(new[] { T("a\tb\nc\r\nd😀e") }));
        Assert.Equal("a\tb\nc\r\nd😀e", back.At(0, 0).Text);
    }

    [Fact]
    public void Lone_surrogates_are_dropped()
    {
        var back = RoundTrip(Grid(new[] { T("a\uD800b\uDC00c") }));
        Assert.Equal("abc", back.At(0, 0).Text);
    }

    [Fact]
    public void Text_that_is_only_control_characters_becomes_an_empty_cell()
    {
        var back = RoundTrip(Grid(new[] { T("\u0000\u0001"), T("x") }));
        Assert.Equal(CellKind.Empty, back.At(0, 0).Kind);
        Assert.Equal("x", back.At(0, 1).Text);
    }

    [Fact]
    public void Control_characters_are_dropped_from_sheet_names()
    {
        Assert.Equal(new[] { "Ab" }, SheetNames(Write(Named("A\u0000b\u0003", new[] { T("x") }))));
        Assert.Equal(new[] { "Sheet1" }, SheetNames(Write(Named("\u0000\u0001", new[] { T("x") }))));
    }

    // ---- stream and file handling ----

    [Fact]
    public void Stream_stays_open_after_WriteXlsx()
    {
        var ms = new MemoryStream();
        Spreadsheet.WriteXlsx(ms, new[] { Grid(new[] { T("x") }) });
        Assert.True(ms.CanWrite);
        ms.WriteByte(0);
    }

    [Fact]
    public void Writes_to_a_non_seekable_stream()
    {
        var inner = new MemoryStream();
        Spreadsheet.WriteXlsx(new ForwardOnlyStream(inner), new[] { Grid(new[] { T("x") }) });
        inner.Position = 0;
        Assert.Equal("x", Spreadsheet.Read(inner, SpreadsheetFormat.Xlsx).At(0, 0).Text);
    }

    [Fact]
    public void WriteFile_xlsx_creates_then_overwrites()
    {
        var path = TempPath(".xlsx");
        try
        {
            Spreadsheet.WriteFile(path, new[] { Grid(new[] { T("first"), T("padding to make the file longer") }) });
            Assert.Equal("first", Spreadsheet.ReadFile(path).At(0, 0).Text);

            Spreadsheet.WriteFile(path, new[] { Grid(new[] { T("second") }) });
            var back = Spreadsheet.ReadFile(path);
            Assert.Equal("second", back.At(0, 0).Text);
            Assert.Equal(1, back.ColumnCount);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void WriteFile_xlsx_writes_all_sheets()
    {
        var path = TempPath(".XLSX");
        try
        {
            Spreadsheet.WriteFile(path, new[] { Named("One", new[] { T("a") }), Named("Two", new[] { T("b") }) });
            Assert.Equal(new[] { "One", "Two" }, SheetNames(File.ReadAllBytes(path)));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void WriteFile_with_unknown_extension_throws_like_the_reader()
    {
        var ex = Assert.Throws<NotSupportedException>(() => Spreadsheet.WriteFile("out.ods", new[] { Grid(new[] { T("x") }) }));
        Assert.Contains("out.ods", ex.Message);
    }

    [Fact]
    public void WriteFile_rejected_sheets_leave_an_existing_file_alone()
    {
        var path = TempPath(".xlsx");
        try
        {
            File.WriteAllText(path, "keep me");
            Assert.Throws<ArgumentException>(() => Spreadsheet.WriteFile(path, new[] { Named(new string('a', 40), new[] { T("x") }) }));
            Assert.Equal("keep me", File.ReadAllText(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void WriteFile_rejects_nulls_and_empty_lists()
    {
        Assert.Throws<ArgumentNullException>(() => Spreadsheet.WriteFile(null!, new[] { Grid() }));
        Assert.Throws<ArgumentNullException>(() => Spreadsheet.WriteFile("a.xlsx", null!));
        Assert.Throws<ArgumentException>(() => Spreadsheet.WriteFile("a.csv", Array.Empty<SheetData>()));
    }

    internal static string TempPath(string ext) => Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ext);

    private sealed class ForwardOnlyStream(Stream inner) : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => inner.Write(buffer, offset, count);
    }
}
