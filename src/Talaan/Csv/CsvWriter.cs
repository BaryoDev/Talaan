using System.Text;

namespace Talaan.Csv;

/// <summary>
/// A small RFC-4180 CSV writer, the mirror of <see cref="CsvReader"/>. A field is quoted when it
/// contains the delimiter, a double quote, CR or LF, and inner quotes are doubled. Cells are written
/// with <see cref="CellValue.AsString"/>, so numbers use invariant culture and dates are ISO-8601.
/// Empty cells become empty fields. Rows end with CRLF, the encoding is UTF-8 with no byte-order mark.
/// </summary>
public static class CsvWriter
{
    /// <summary>
    /// Writes <paramref name="sheet"/> to <paramref name="stream"/> as CSV (or TSV, with a tab
    /// <paramref name="delimiter"/>). The stream is left open.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="stream"/> or <paramref name="sheet"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="delimiter"/> is a quote, CR or LF.</exception>
    public static void Write(Stream stream, SheetData sheet, char delimiter = ',')
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(sheet);
        if (delimiter is '"' or '\r' or '\n')
            throw new ArgumentException("The delimiter cannot be a quote, CR or LF.", nameof(delimiter));

        using var writer = new StreamWriter(stream, new UTF8Encoding(false), bufferSize: 4096, leaveOpen: true)
        {
            NewLine = "\r\n"
        };
        foreach (var row in sheet.Rows)
        {
            for (var c = 0; c < row.Count; c++)
            {
                if (c > 0) writer.Write(delimiter);
                WriteField(writer, row[c].AsString(), delimiter);
            }
            writer.WriteLine();
        }
    }

    private static void WriteField(TextWriter writer, string value, char delimiter)
    {
        if (value.IndexOfAny(new[] { delimiter, '"', '\r', '\n' }) < 0)
        {
            writer.Write(value);
            return;
        }
        writer.Write('"');
        writer.Write(value.Replace("\"", "\"\""));
        writer.Write('"');
    }
}
