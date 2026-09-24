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
