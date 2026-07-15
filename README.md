<div align="center">
  <img src="assets/logo.svg" width="96" height="96" alt="Talaan logo" />
  <h1>Talaan</h1>
  <p><em>A zero-dependency spreadsheet &amp; CSV reader for .NET.</em></p>
</div>

---

**Talaan** (Filipino for *record / list / table*) reads `.xlsx` and CSV files into a simple, typed
cell grid — with **no external dependencies**. No Excel Interop, no ClosedXML, no OpenXML SDK. An
`.xlsx` is just a zip of XML, and Talaan walks those parts directly using only the .NET base class
library (`System.IO.Compression` + `System.Xml`).

## Why

- **Zero dependencies.** Nothing to audit, nothing to update, no transitive surprises. One small assembly.
- **Typed cells.** Numbers come back as numbers, dates as `DateTime`, booleans as `bool` — not stringly-typed guesses. Excel serial dates and shared strings are resolved for you.
- **Raw grid, no magic.** Talaan hands you the sheet exactly as laid out. *You* decide which row is the header and which rows to skip — perfect for real-world files with title rows, sections, and blanks.

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
| Excel  | `.xlsx`   | First worksheet. Shared strings, inline strings, booleans, numbers, and style-based date detection. |
| CSV    | `.csv`, `.txt` | RFC-4180: quoted fields, escaped `""`, embedded commas and newlines. |
| TSV    | `.tsv`    | Tab-delimited. |

You can also bypass detection: `Spreadsheet.Read(stream, SpreadsheetFormat.Xlsx)`.

## Design notes

- **Dates.** `.xlsx` stores dates as serial numbers; Talaan inspects the workbook stylesheet
  (`cellXfs` → number format) to decide whether a numeric cell is a date, then converts via the OLE
  Automation epoch. Ambiguous `m` (month vs. minute) is ignored in favour of unambiguous `y`/`d`.
- **Column alignment.** Cell references (`A1`, `AB12`) are honoured, so skipped/blank cells become
  `CellValue.Empty` and columns stay aligned.
- **Streaming input.** Non-seekable streams are buffered automatically for zip access.

## Scope (by design)

Talaan **reads**; it does not write. It reads the **first worksheet**. Formulas are returned by their
last cached result. If you need multi-sheet, writing, or styling, this isn't that library — and that
is the point.

## License

[MPL-2.0](LICENSE) © BaryoDev
