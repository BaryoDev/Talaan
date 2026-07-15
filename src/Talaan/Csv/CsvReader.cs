using System.Text;

namespace Talaan.Csv;

/// <summary>
/// A small RFC-4180 CSV reader. Handles quoted fields, escaped quotes (""), and commas/newlines
/// inside quotes. Every field is returned as a <see cref="CellKind.Text"/> cell (CSV carries no
/// type information); empty fields become <see cref="CellValue.Empty"/>.
/// </summary>
public static class CsvReader
{
    public static SheetData Read(Stream stream, char delimiter = ',')
    {
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
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

        // Flush a trailing field/row only if the file didn't end on a clean newline.
        if (fieldHasContent || field.Length > 0 || row.Count > 0 || rowHasContent)
            EndRow();

        return new SheetData(rows);
    }
}
