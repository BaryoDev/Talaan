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
    public static SheetData Read(Stream stream, string fileName)
        => Read(stream, DetectFormat(fileName));

    public static SheetData Read(Stream stream, SpreadsheetFormat format) => format switch
    {
        SpreadsheetFormat.Xlsx => XlsxReader.Read(stream),
        SpreadsheetFormat.Tsv => CsvReader.Read(stream, '\t'),
        SpreadsheetFormat.Csv => CsvReader.Read(stream, ','),
        _ => throw new NotSupportedException($"Unsupported spreadsheet format: {format}.")
    };

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
