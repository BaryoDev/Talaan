using System.Text;
using System.Xml;
using Talaan.Csv;
using Talaan.Xlsx;

namespace Talaan;

/// <summary>
/// Entry point for reading a spreadsheet into a <see cref="SheetData"/> grid. Dispatches by file
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
