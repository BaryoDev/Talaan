# Contributing

Contributions are welcome, including small ones. This adds to the
[BaryoDev-wide guide](https://github.com/BaryoDev/.github/blob/main/CONTRIBUTING.md); where the two
disagree, this file wins.

## Getting it running

```sh
git clone https://github.com/BaryoDev/Talaan.git
cd Talaan
dotnet test
```

That is the whole setup. No database, no fixture download, no local paths. The suite runs in about
a second, and if it takes longer than that something is wrong.

The library targets `net8.0` and has no package references at all. The test project references
xunit and nothing else.

## What this library is for

One thing, and it decides most design questions: **read a spreadsheet without taking on a
dependency.** No Interop, no ClosedXML, no OpenXML SDK. An `.xlsx` is a zip of XML parts and Talaan
walks those parts with `System.IO.Compression` and `System.Xml.Linq`, both of which are in the base
class library.

That makes the dependency list a hard constraint rather than a preference. **A pull request that
adds a `PackageReference` to `src/Talaan/Talaan.csproj` is a pull request that removes the reason
this library exists.** If you hit something that seems to need a package, open an issue first: the
answer is usually that we write the twenty lines, or that we decide the feature is out of scope.

Second constraint, from `ROADMAP.md`: the published `0.1.0` surface (`Spreadsheet.Read*`,
`SheetData`, `CellValue`) does not change. New capability arrives as new types and new methods, not
as changed signatures.

## Things worth knowing before you change something

**Talaan reads the first worksheet only, and that is deliberate for now.** `XlsxReader.Read` calls
`ResolveFirstSheetPath`, which maps the first `<sheet>` in `xl/workbook.xml` through
`xl/_rels/workbook.xml.rels` to a part path, and falls back to the alphabetically first part under
`xl/worksheets/` if that lookup fails. Multi-sheet is planned as an additive `Workbook` type
(`ROADMAP.md`, v0.2.0), not as a changed return type.

**Dates are the hardest part of this format and the code is opinionated about it.** `.xlsx` stores
a date as a plain number; whether it *is* a date lives in the stylesheet. `ReadDateStyles` walks
`cellXfs` in document order, because a cell's `s="4"` attribute is an index into that order, and
marks the indices whose `numFmtId` is either a built-in date format or a custom one whose format
code looks like a date. `LooksLikeDate` strips quoted literals and `[...]` sections, then keys off
`y` and `d` only. **`m` is ignored on purpose**, because it means month in `mm/dd` and minute in
`hh:mm` and there is no way to tell from one character. The consequence, which is a real tradeoff
and not an oversight, is that a time-only custom format such as `hh:mm:ss` is not detected as a
date. If you change `LooksLikeDate`, you are trading false positives against false negatives, so
say in the PR which direction you chose and why.

**Column gaps are padded. Row gaps currently are not.** Excel omits cells and whole rows that have
no content, so `<row r="3">` can follow `<row r="1">` with nothing between. The cell loop honours
the `r` attribute on `<c>` and pads with `CellValue.Empty` to keep columns aligned. The row loop
does not do the same with the `r` attribute on `<row>`, so a blank row collapses the grid. That is
a known bug with an open issue, not a design decision, and it is worth knowing about before you
write a test that happens to encode the current behaviour.

**Shared strings are concatenated from every `<t>` under an `<si>`.** That is correct for rich text,
where one string is split across several `<r><t>` runs and has to be rejoined. It is wrong for
phonetic guides (`<rPh>`), which are also `<t>` elements. Known bug, open issue.

**Formulas are not evaluated.** A formula cell is returned by its cached `<v>` result, which is
whatever Excel last wrote. A file saved by a tool that does not write cached values reads as empty,
and that is expected.

**Talaan does not write.** Not to `.xlsx`, and not yet to CSV (a CSV writer is `ROADMAP.md` v0.3.0).

## Fixtures live in memory, not in the repository

There are no `.xlsx` or `.csv` files committed to this repository and preferably there never will
be. A binary fixture is unreviewable in a diff, tends to carry whatever was in the file the author
happened to have, and the last one deleted from this repo was asserting against real personal data
from a hard-coded local path.

CSV tests build their input from a string literal, which you can read in the test:

```csharp
private static SheetData Parse(string csv)
    => CsvReader.Read(new MemoryStream(Encoding.UTF8.GetBytes(csv)));
```

The xlsx equivalent is a small helper that zips up the OOXML parts you care about. If it does not
exist yet when you read this, building it is the open issue to pick up first, and every other xlsx
issue is waiting behind it. The point of both is the same: a reader of your test can see the exact
bytes that go in, so the test documents the format quirk it is pinning.

## Tests

**Every change needs a test that fails without it.** If you cannot write one, say so in the
description and explain why.

The worked version of that rule, with before-and-after examples, is at
[BaryoDev/.github/TESTING.md](https://github.com/BaryoDev/.github/blob/main/TESTING.md). The
question it turns on is worth asking yourself before you push:

> Name the production change that would make this test fail.

If the answer is "deleting the method" or "returning null", the test is pinning the wiring, not the
behaviour. For a file-format reader there is almost always a sharper test available, because the
input is bytes and the output is a grid: a malformed file, a specific encoding, a date format, a
shared-string edge case, a stream that ends early. Each of those is a fixture and an assertion, and
each one fails loudly when the parsing changes.

A concrete example of the difference, on the date path:

```csharp
// Pins the wiring. Passes if FromOADate is called on the wrong epoch,
// if the style lookup is off by one, or if every cell is a date.
Assert.Equal(CellKind.Date, sheet.At(0, 0).Kind);

// Pins the behaviour. Fails if the epoch, the style index or the
// format detection is wrong, and names the value it expects.
Assert.Equal(new DateTime(2026, 9, 1), sheet.At(0, 0).Date);
```

Bug fixes get the same treatment, and it is easier there: write the test that reproduces the bug,
watch it fail, then fix it. A bug fix whose test passes before the fix has not reproduced the bug.

CI runs `dotnet test` and builds with warnings as errors on every pull request. It has to pass
before a merge.

## Good first issues

Look for [`good first issue`](https://github.com/BaryoDev/Talaan/labels/good%20first%20issue). Most
of them are one function and one test, and several are format quirks with a known correct answer
written into the issue, so you are not guessing at intent.

If nothing fits, an issue describing the file that Talaan read wrong, with the smallest sheet that
reproduces it, is genuinely useful. This library gets better mostly by meeting files it has not met.

## One entirely optional favour

If this library saved you an afternoon, a star costs you nothing and is the only signal a small
free package has that anyone is using it. It is not a condition of anything: your pull request gets
the same review either way, and nobody is checking.

## Reporting a security issue

Do not open a public issue. See [SECURITY.md](SECURITY.md).
```

---
