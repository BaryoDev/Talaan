<div align="center">
  <img src="https://raw.githubusercontent.com/BaryoDev/Talaan/master/assets/icon.png" width="96" height="96" alt="Talaan logo" />
  <h1>Talaan</h1>
  <p><em>A zero-dependency spreadsheet &amp; CSV reader for .NET.</em></p>
</div>

---

**Talaan** (Filipino for *record / list / table*) reads and writes `.xlsx` and CSV files as a simple, typed
cell grid, with **no external dependencies**. No Excel Interop, no ClosedXML, no OpenXML SDK. An
`.xlsx` is just a zip of XML, and Talaan walks those parts directly using only the .NET base class
library (`System.IO.Compression` + `System.Xml`).

## Why

- **Zero dependencies.** Nothing to audit, nothing to update, no transitive surprises. One small assembly.
- **Typed cells.** Numbers come back as numbers, dates as `DateTime`, booleans as `bool`, not stringly-typed guesses. Excel serial dates and shared strings are resolved for you.
- **Raw grid, no magic.** Talaan hands you the sheet exactly as laid out. *You* decide which row is the header and which rows to skip, which suits real-world files with title rows, sections, and blanks.

## Install

```bash
dotnet add package Talaan
```

## Quick start

```csharp
using Talaan;

// From a file…
SheetData sheet = Spreadsheet.ReadFile("members.xlsx");

// …or a stream (an upload, say), dispatched by file name:
SheetData sheet = Spreadsheet.Read(uploadStream, fileName);

for (int r = 0; r < sheet.RowCount; r++)
{
    CellValue name = sheet.At(r, 1);   // row r, column B
    if (name.Kind == CellKind.Text)
        Console.WriteLine(name.Text);
}
```

Every cell is a `CellValue`:

```csharp
public readonly record struct CellValue(
    CellKind Kind,        // Empty | Text | Number | Date | Boolean
    string?  Text,
    double?  Number,
    DateTime? Date,
    bool?    Boolean);

cell.IsBlank      // empty, or whitespace-only text
cell.AsString()   // best-effort display: invariant numbers, ISO-8601 dates
```

## Supported formats

| Format | Extension | Notes |
|--------|-----------|-------|
| Excel  | `.xlsx`   | Read: first worksheet. Write: every sheet. Shared strings, inline strings, booleans, numbers, and style-based date detection. |
| CSV    | `.csv`, `.txt` | RFC-4180: quoted fields, escaped `""`, embedded commas and newlines. |
| TSV    | `.tsv`    | Tab-delimited. |

You can also bypass detection: `Spreadsheet.Read(stream, SpreadsheetFormat.Xlsx)`.

### Reading a CSV in a non-UTF-8 encoding

CSV defaults to UTF-8. A file exported by "Excel on Windows > CSV (Comma delimited)" is usually
Windows-1252 with no byte-order mark, so it needs an explicit encoding:

```csharp
using System.Text;

// Code page 1252 needs the CodePages package, which Talaan does not depend on:
//   dotnet add package System.Text.Encoding.CodePages
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
SheetData sheet = Spreadsheet.ReadFile("members.csv", Encoding.GetEncoding(1252));
```

`Encoding.Latin1` is built in and works without the package, but it is not the same as 1252:
bytes 0x80 to 0x9F (the euro sign, curly quotes, dashes) come back as control characters.

`CsvReader.Read(stream, encoding)` takes the same encoding for CSV/TSV without going through
`Spreadsheet`.

## Writing

```csharp
var people = new SheetData(new IReadOnlyList<CellValue>[]
{
    new[] { CellValue.OfText("name"), CellValue.OfText("joined"), CellValue.OfText("active") },
    new[] { CellValue.OfText("Ana"), CellValue.OfDate(new DateTime(2024, 5, 6)), CellValue.OfBoolean(true) },
}, name: "People");

// To a file. The extension picks the format and an existing file is overwritten.
Spreadsheet.WriteFile("members.xlsx", new[] { people, otherSheet });  // every sheet
Spreadsheet.WriteFile("members.csv", new[] { people });               // first sheet only
Spreadsheet.WriteFile("members.tsv", new[] { people });               // first sheet only, tab-delimited

// To a stream. The stream is left open.
Spreadsheet.WriteXlsx(responseStream, new[] { people });
CsvWriter.Write(responseStream, people, delimiter: ',');             // Talaan.Csv
```

- **Text** is written as inline strings, so `"00123"` stays text. Booleans and numbers are typed
  cells. Numbers use invariant culture and round-trip exactly. NaN and infinity throw.
- **Dates** are written as serial numbers. A date with no time gets the number format
  `yyyy-mm-dd`, one with a time gets `yyyy-mm-dd hh:mm:ss`. Both read back as `CellKind.Date`.
  Dates before 1900-03-01 throw, because Excel's serials for them differ from .NET's by a day.
- **Empty cells** are skipped. Cell references stay correct past column Z, up to XFD.
- **Characters XML 1.0 forbids** (most control characters, lone surrogates) are dropped from text
  and sheet names.
- **Sheet names**: a null or blank name becomes `Sheet{N}` (its 1-based position). A name over 31
  characters, containing any of `[ ] : * ? / \`, starting or ending with an apostrophe, or repeating
  another name (case-insensitive) throws `ArgumentException`. An empty sheet list also throws.
- **CSV and TSV** follow RFC 4180: a field is quoted when it holds the delimiter, a quote, CR or
  LF, and inner quotes are doubled. Rows end with CRLF. UTF-8, no byte-order mark. Cells are written
  with `AsString()`.

## Design notes

- **Dates.** `.xlsx` stores dates as serial numbers; Talaan inspects the workbook stylesheet
  (`cellXfs` to number format) to decide whether a numeric cell is a date, then converts via the OLE
  Automation epoch. Ambiguous `m` (month vs. minute) is ignored in favour of unambiguous `y`/`d`.
- **Column alignment.** Cell references (`A1`, `AB12`) are honoured, so skipped/blank cells become
  `CellValue.Empty` and columns stay aligned.
- **Streaming input.** Non-seekable streams are buffered automatically for zip access.

## Scope (by design)

Reading returns the **first worksheet** only. Formulas are returned by their last cached result.
Writing produces a plain workbook: values and the date formats above, no styling, formulas, column
widths or merged cells. If you need those, this isn't that library, and that is the point.

## License

[MPL-2.0](LICENSE) © BaryoDev
