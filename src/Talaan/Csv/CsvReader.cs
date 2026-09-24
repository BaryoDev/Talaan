using System.Text;

namespace Talaan.Csv;

/// <summary>
/// A small RFC-4180 CSV reader. Handles quoted fields, escaped quotes (""), and commas/newlines
/// inside quotes. Every field is returned as a <see cref="CellKind.Text"/> cell (CSV carries no
/// type information); empty fields become <see cref="CellValue.Empty"/>.
///
/// A quoted field left open at end of stream throws <see cref="InvalidDataException"/> naming the
/// record it started in, rather than silently swallowing the rest of the file: returning one row
/// where the file had ten thousand is data loss, not leniency. Text found right after a closing
/// quote (<c>"a"b</c>) is still appended to the same field instead of rejected; that loses no data,
/// so it stays lenient for now.
/// </summary>
public static class CsvReader
{
    public static SheetData Read(Stream stream, char delimiter = ',')
    {
        using var reader = new StreamReader(
            stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: -1, leaveOpen: true);
        return Read(reader, delimiter);
    }

    /// <summary>
    /// Reads CSV/TSV using <paramref name="encoding"/> instead of the UTF-8 default, for files such
    /// as a Windows-1252 export from Excel on Windows. A byte-order mark, if present, still wins:
    /// this overload keeps <c>detectEncodingFromByteOrderMarks: true</c>, matching
    /// <see cref="Read(Stream, char)"/>, so a UTF-8 or UTF-16 BOM overrides <paramref name="encoding"/>
    /// rather than being misread as data.
    /// <para>
    /// A single-byte Windows code page such as 1252 needs <c>Encoding.GetEncoding(1252)</c>, which
    /// throws on .NET 8 unless the caller has registered
    /// <c>System.Text.Encoding.CodePages</c>'s <c>CodePagesEncodingProvider</c>. Talaan does not take
    /// that package as a dependency, so that registration is the caller's job.
    /// </para>
    /// </summary>
    public static SheetData Read(Stream stream, Encoding encoding, char delimiter = ',')
    {
        using var reader = new StreamReader(
            stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: -1, leaveOpen: true);
        return Read(reader, delimiter);
    }

    public static SheetData Read(TextReader reader, char delimiter = ',')
    {
        var rows = new List<IReadOnlyList<CellValue>>();
        var row = new List<CellValue>();
        var field = new StringBuilder();
        bool inQuotes = false;
        bool fieldHasContent = false; // distinguishes a started field from a brand-new row
        bool rowHasContent = false;
        int quoteStartRecord = 0; // 1-based record the currently open quote began in, 0 when not in a quote

        void EndField()
        {
            row.Add(field.Length == 0 ? CellValue.Empty : CellValue.OfText(field.ToString()));
            field.Clear();
            fieldHasContent = false;
        }

        void EndRow()
        {
            EndField();
            rows.Add(row.ToArray());
            row = new List<CellValue>();
            rowHasContent = false;
        }

        int read;
        while ((read = reader.Read()) != -1)
        {
            var c = (char)read;

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (reader.Peek() == '"') { field.Append('"'); reader.Read(); } // escaped quote
                    else inQuotes = false;
                }
                else
                {
                    field.Append(c);
                }
                continue;
            }

            if (c == '"')
            {
                inQuotes = true;
                quoteStartRecord = rows.Count + 1;
                fieldHasContent = true;
                rowHasContent = true;
            }
            else if (c == delimiter)
            {
                EndField();
                rowHasContent = true;
            }
            else if (c == '\r')
            {
                if (reader.Peek() == '\n') reader.Read();
                EndRow();
            }
            else if (c == '\n')
            {
                EndRow();
            }
            else
            {
                field.Append(c);
                fieldHasContent = true;
                rowHasContent = true;
            }
        }

        if (inQuotes)
            throw new InvalidDataException($"Unterminated quoted field: record {quoteStartRecord} opens a quote that is never closed.");

        // Flush a trailing field/row only if the file didn't end on a clean newline.
        if (fieldHasContent || field.Length > 0 || row.Count > 0 || rowHasContent)
            EndRow();

        return new SheetData(rows);
    }
}
