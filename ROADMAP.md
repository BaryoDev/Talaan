# Talaan roadmap

Incremental enhancements. Each version is a small, independently releasable step, and **nothing
breaks the published `0.1.0` API** (`Spreadsheet.Read*`, `SheetData`, `CellValue` stay as-is).

Guiding rules: stay **zero-dependency**, keep reading the common case a one-liner, and add new
surface additively (new types/methods, never changed signatures).

---

## v0.2.0 — Read different sheets (multi-sheet)

Today `XlsxReader` reads only the first worksheet. Add whole-workbook access without touching the
existing one-sheet entry points.

**New type:**
```csharp
public sealed class Workbook
{
    public IReadOnlyList<SheetData> Sheets { get; }
    public IReadOnlyList<string> SheetNames { get; }   // in document order
    public SheetData this[int index] { get; }
    public SheetData this[string name] { get; }        // case-insensitive; throws if missing
    public bool TryGet(string name, out SheetData sheet);
}
```

**New entry points (additive):**
```csharp
Workbook wb = Spreadsheet.ReadWorkbook(stream, fileName);   // xlsx: all sheets; csv: one sheet
Workbook wb = XlsxReader.ReadWorkbook(stream);
SheetData s = Spreadsheet.ReadSheet(stream, fileName, "Members");   // by name
SheetData s = Spreadsheet.ReadSheet(stream, fileName, index: 1);    // by index
```

**Existing behavior preserved:** `Spreadsheet.Read(...)` keeps returning the **first** sheet
(implemented as `ReadWorkbook(...).Sheets[0]`). `SheetData.Name` — already on the type but currently
null — gets **populated** with the worksheet name (additive, safe).

**Implementation notes (small):**
- Refactor the existing `ResolveFirstSheetPath` into `EnumerateSheets` — parse `xl/workbook.xml`
  `<sheet name= r:id=>` + `xl/_rels/workbook.xml.rels` to get **(name, part path)** for every sheet,
  in order. The first-sheet path already comes from here, so this is mostly generalizing what exists.
- Read each sheet part with the current cell-parsing loop; shared strings + styles are workbook-wide,
  so parse them once and reuse across sheets.
- CSV has one sheet: `ReadWorkbook` wraps it, named after the file (sans extension).

Effort: small. One new type + a loop around existing parsing.

---

## v0.3.0 — Write CSV

Talaan is read-only today. Add a CSV **writer** (mirror of the RFC-4180 reader).

**New API (additive):**
```csharp
CsvWriter.Write(SheetData sheet, TextWriter writer, char delimiter = ',');
CsvWriter.Write(SheetData sheet, Stream stream, char delimiter = ',');
Spreadsheet.WriteCsv(SheetData sheet, string path);
string csv = sheet.ToCsv();                 // convenience extension
```

**Also write from arbitrary rows** (not just a round-tripped `SheetData`):
```csharp
CsvWriter.Write(IEnumerable<IEnumerable<object?>> rows, TextWriter writer, char delimiter = ',');
```

**Correctness (the whole point):**
- Quote a field when it contains the delimiter, `"`, `\r`, or `\n`; escape `"` as `""`.
- `CellValue` serializes via its existing `AsString()` (invariant numbers, ISO dates), so read→write
  round-trips cleanly.
- Configurable newline (default `\n`); UTF-8, optional BOM flag for Excel-friendliness.

Effort: small. A focused state-free writer + a couple of convenience overloads.

---

## Later (parked, only if needed)

- **v0.4.0 — Header-aware records.** `sheet.AsRecords(headerRow: 0)` → `IEnumerable<IReadOnlyDictionary<string, CellValue>>`, keyed by header. Optional `MapTo<T>()` by property-name match. Turns the raw grid into typed data (useful for the BaryoClub import flow).
- **v0.5.0 — Streaming reads.** `XlsxReader.ReadRows(stream)` / `CsvReader.ReadRows(...)` returning `IEnumerable<IReadOnlyList<CellValue>>` so large files don't load fully into memory. Additive; the current eager `Read` stays.
- **TSV / custom-delimiter write** — trivial once the CSV writer exists (delimiter param already there).
- **XLSX writing** — the big one (author the OOXML zip: shared strings, styles, sheet XML). Meaningfully larger than everything above; only if there's real demand. Keep it a separate `Talaan.Xlsx.Writer` so the read path stays lean.

---

## Sequencing

1. **v0.2.0** multi-sheet read — highest value, smallest change, reuses existing parsing.
2. **v0.3.0** CSV write — small, self-contained, and unlocks read-xlsx → write-csv pipelines.
3. Header records / streaming as real use cases appear.
4. XLSX write only if demanded.

Each ships via the existing `Publish to NuGet` workflow with a bumped `Version`.
