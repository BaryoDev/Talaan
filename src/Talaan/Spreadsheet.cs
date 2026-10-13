using System.Text;
using System.Xml;
using Talaan.Csv;
using Talaan.Xlsx;

namespace Talaan;

/// <summary>
/// Entry point for reading a spreadsheet into a <see cref="SheetData"/> grid, and for writing
/// sheets back out as .xlsx, .csv or .tsv. Dispatches by file
/// extension (or an explicit <see cref="SpreadsheetFormat"/>): .xlsx via <see cref="XlsxReader"/>,
/// .csv/.tsv/.txt via <see cref="CsvReader"/>.
/// </summary>
public static class Spreadsheet
{
    /// <summary>Reads a spreadsheet, detecting the format from <paramref name="fileName"/>'s extension.</summary>
    /// <exception cref="InvalidDataException">The stream is not a valid file of the detected format:
    /// not a valid zip for .xlsx, a part that is not well-formed XML or carries a DOCTYPE (the
    /// message names the part and <see cref="Exception.InnerException"/> is the original
    /// <see cref="XmlException"/>), or an unterminated quoted field for CSV/TSV.</exception>
    public static SheetData Read(Stream stream, string fileName)
        => Read(stream, DetectFormat(fileName));

    /// <summary>Reads a spreadsheet in the given format.</summary>
    /// <exception cref="InvalidDataException">The stream is not a valid file of
    /// <paramref name="format"/>: not a valid zip for .xlsx, a part that is not well-formed XML or
    /// carries a DOCTYPE (the message names the part and <see cref="Exception.InnerException"/> is
    /// the original <see cref="XmlException"/>), or an unterminated quoted field for CSV/TSV.</exception>
    public static SheetData Read(Stream stream, SpreadsheetFormat format) => format switch
    {
        SpreadsheetFormat.Xlsx => XlsxReader.Read(stream),
        SpreadsheetFormat.Tsv => CsvReader.Read(stream, '\t'),
        SpreadsheetFormat.Csv => CsvReader.Read(stream, ','),
        _ => throw new NotSupportedException($"Unsupported spreadsheet format: {format}.")
    };

    /// <summary>
    /// Reads a spreadsheet in the given format, using <paramref name="encoding"/> to decode CSV/TSV
    /// text instead of the UTF-8 default. Ignored when <paramref name="format"/> is
    /// <see cref="SpreadsheetFormat.Xlsx"/>, since an xlsx part declares its own XML encoding.
    /// See <see cref="Csv.CsvReader.Read(Stream, Encoding, char)"/> for how a byte-order mark
    /// interacts with <paramref name="encoding"/>, and for the Windows code page note.
    /// </summary>
    /// <exception cref="InvalidDataException">The stream is not a valid file of
    /// <paramref name="format"/> (not a valid zip for .xlsx, or an unterminated quoted field for
    /// CSV/TSV).</exception>
    /// <exception cref="XmlException">The format is <see cref="SpreadsheetFormat.Xlsx"/> and a
    /// part is not well-formed XML, or carries a DOCTYPE. DTDs are never processed; this is not
    /// wrapped in a different exception type.</exception>
    public static SheetData Read(Stream stream, SpreadsheetFormat format, Encoding encoding) => format switch
    {
        SpreadsheetFormat.Xlsx => XlsxReader.Read(stream),
        SpreadsheetFormat.Tsv => CsvReader.Read(stream, encoding, '\t'),
        SpreadsheetFormat.Csv => CsvReader.Read(stream, encoding, ','),
        _ => throw new NotSupportedException($"Unsupported spreadsheet format: {format}.")
    };

    /// <summary>Reads a spreadsheet from disk, detecting the format from the file extension.</summary>
    /// <exception cref="InvalidDataException">The file is not a valid file of the detected format:
    /// not a valid zip for .xlsx, a part that is not well-formed XML or carries a DOCTYPE (the
    /// message names the part and <see cref="Exception.InnerException"/> is the original
    /// <see cref="XmlException"/>), or an unterminated quoted field for CSV/TSV.</exception>
    public static SheetData ReadFile(string path)
    {
        using var fs = File.OpenRead(path);
        return Read(fs, DetectFormat(path));
    }

    /// <summary>
    /// Reads a spreadsheet from disk, detecting the format from the file extension and using
    /// <paramref name="encoding"/> to decode CSV/TSV text instead of the UTF-8 default. Ignored for
    /// .xlsx files. See <see cref="Csv.CsvReader.Read(Stream, Encoding, char)"/> for the byte-order
    /// mark behavior and the Windows code page note.
    /// </summary>
    /// <exception cref="InvalidDataException">The file is not a valid file of the detected format
    /// (not a valid zip for .xlsx, or an unterminated quoted field for CSV/TSV).</exception>
    /// <exception cref="XmlException">The file is .xlsx and a part is not well-formed XML, or
    /// carries a DOCTYPE. DTDs are never processed; this is not wrapped in a different exception
    /// type.</exception>
    public static SheetData ReadFile(string path, Encoding encoding)
    {
        using var fs = File.OpenRead(path);
        return Read(fs, DetectFormat(path), encoding);
    }

    /// <summary>
    /// Writes <paramref name="sheets"/> as an .xlsx workbook, one worksheet per sheet in order. The
    /// stream is left open. Text is written as inline strings, booleans and numbers as typed cells,
    /// and dates as serial numbers styled <c>yyyy-mm-dd</c> (or <c>yyyy-mm-dd hh:mm:ss</c> when the
    /// value has a time), so <see cref="Read(Stream, SpreadsheetFormat)"/> reads them back as
    /// <see cref="CellKind.Date"/>. Empty cells are skipped. Characters XML 1.0 does not allow are
    /// dropped from text and sheet names.
    /// <para>
    /// A sheet with a null or blank <see cref="SheetData.Name"/> is named <c>Sheet{N}</c>, where N is
    /// its 1-based position.
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="stream"/> or <paramref name="sheets"/> is null.</exception>
    /// <exception cref="ArgumentException">The list is empty, or a sheet name is longer than 31
    /// characters, contains one of <c>[ ] : * ? / \</c>, starts or ends with an apostrophe, or repeats
    /// another name (compared without regard to case). Also thrown for a number that is NaN or
    /// infinite, a date before 1900-03-01, or a sheet larger than 1048576 rows by 16384 columns.</exception>
    public static void WriteXlsx(Stream stream, IEnumerable<SheetData> sheets)
        => XlsxWriter.Write(stream, sheets);

    /// <summary>
    /// Writes a spreadsheet to disk, creating the file or overwriting it. The format comes from the
    /// extension: <c>.xlsx</c> writes every sheet, <c>.csv</c> and <c>.txt</c> (comma) and <c>.tsv</c>
    /// (tab) write the first sheet only. Nothing is written to disk if the sheets are rejected.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> or <paramref name="sheets"/> is null.</exception>
    /// <exception cref="ArgumentException">The list is empty, or the sheets are rejected as described
    /// on <see cref="WriteXlsx(Stream, IEnumerable{SheetData})"/>.</exception>
    /// <exception cref="NotSupportedException">The extension is not .xlsx, .csv, .txt or .tsv.</exception>
    public static void WriteFile(string path, IEnumerable<SheetData> sheets)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentNullException.ThrowIfNull(sheets);
        var format = DetectFormat(path);
        var list = sheets.ToList();
        if (list.Count == 0)
            throw new ArgumentException("At least one sheet is required.", nameof(sheets));
        if (list[0] is null)
            throw new ArgumentException("The sheet list contains a null sheet.", nameof(sheets));

        using var buffer = new MemoryStream();
        switch (format)
        {
            case SpreadsheetFormat.Xlsx: WriteXlsx(buffer, list); break;
            case SpreadsheetFormat.Tsv: CsvWriter.Write(buffer, list[0], '\t'); break;
            case SpreadsheetFormat.Csv: CsvWriter.Write(buffer, list[0], ','); break;
            default: throw new NotSupportedException($"Unsupported spreadsheet format: {format}.");
        }
        File.WriteAllBytes(path, buffer.ToArray());
    }

    public static SpreadsheetFormat DetectFormat(string fileName)
    {
        var ext = Path.GetExtension(fileName).ToLowerInvariant();
        return ext switch
        {
            ".xlsx" => SpreadsheetFormat.Xlsx,
            ".tsv" => SpreadsheetFormat.Tsv,
            ".csv" or ".txt" => SpreadsheetFormat.Csv,
            _ => throw new NotSupportedException($"Cannot infer spreadsheet format from '{fileName}'.")
        };
    }
}

public enum SpreadsheetFormat
{
    Csv,
    Tsv,
    Xlsx
}
