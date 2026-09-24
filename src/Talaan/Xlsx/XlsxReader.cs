using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace Talaan.Xlsx;

/// <summary>
/// A zero-dependency reader for the first worksheet of an .xlsx (OOXML) workbook. An .xlsx file is a
/// zip of XML parts; this walks those parts directly using <see cref="ZipArchive"/> and LINQ-to-XML —
/// no Interop, ClosedXML, or OpenXML SDK. Resolves shared strings, detects date-formatted cells via
/// the stylesheet, and maps cell references so gaps become empty cells.
/// </summary>
public static class XlsxReader
{
    // Built-in number-format ids that denote dates/times (per the OOXML spec).
    private static readonly HashSet<int> BuiltinDateFormats =
        new(new[] { 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47 });

    // A workbook part never legitimately has a DOCTYPE. Prohibiting it also blocks entity
    // expansion (billion-laughs style DoS), which is the actual risk: external entities already
    // fail to resolve with no XmlResolver, but an internal DTD subset is still parsed by default.
    // IgnoreWhitespace = true matches what plain unguarded XML loading did before this fix: a
    // <t> with no xml:space="preserve" and only whitespace content reads as "", one with
    // xml:space="preserve" still reads as the literal whitespace. This is a security fix, not
    // the place to change that.
    private static readonly XmlReaderSettings PartReaderSettings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreWhitespace = true,
    };

    /// <summary>Loads an XML part with DTDs prohibited. Every part Talaan reads must go through this.</summary>
    private static XDocument LoadPart(Stream stream)
    {
        using var reader = XmlReader.Create(stream, PartReaderSettings);
        return XDocument.Load(reader);
    }

    // Excel's row limit (2^20). A row reference above this can't be a real file, and letting it
    // through would pad an unbounded number of empty rows.
    private const int MaxRowNumber = 1_048_576;

    /// <summary>Reads the first worksheet of an .xlsx stream into a <see cref="SheetData"/> grid.</summary>
    /// <exception cref="InvalidDataException">The stream is not a valid zip, or a required part
    /// (the worksheet itself) is missing.</exception>
    /// <exception cref="XmlException">A part is not well-formed XML, or carries a DOCTYPE.
    /// DTDs are never processed; this is not wrapped in a different exception type.</exception>
    public static SheetData Read(Stream stream)
    {
        // ZipArchive needs a seekable stream; buffer if necessary.
        Stream seekable = stream.CanSeek ? stream : Buffer(stream);
        using var archive = new ZipArchive(seekable, ZipArchiveMode.Read, leaveOpen: true);

        var sharedStrings = ReadSharedStrings(archive);
        var dateStyles = ReadDateStyles(archive);
        var date1904 = ReadDate1904(archive);
        var sheetPath = ResolveFirstSheetPath(archive);

        var entry = GetEntry(archive, sheetPath)
            ?? throw new InvalidDataException($"Worksheet part '{sheetPath}' not found in workbook.");

        using var sheetStream = entry.Open();
        var doc = LoadPart(sheetStream);

        var rows = new List<IReadOnlyList<CellValue>>();
        foreach (var rowEl in Descendants(doc.Root, "row"))
        {
            var rowIndex = RowIndex((string?)rowEl.Attribute("r"));
            // Pad any skipped rows with an empty row, same idea as the column padding below.
            while (rowIndex >= 0 && rows.Count < rowIndex) rows.Add(Array.Empty<CellValue>());

            var cells = new List<CellValue>();
            foreach (var cEl in Elements(rowEl, "c"))
            {
                var colIndex = ColumnIndex((string?)cEl.Attribute("r"));
                // Pad any skipped columns with empties so alignment is preserved.
                while (colIndex >= 0 && cells.Count < colIndex) cells.Add(CellValue.Empty);
                cells.Add(ParseCell(cEl, sharedStrings, dateStyles, date1904));
            }
            rows.Add(cells);
        }

        return new SheetData(rows, name: null);
    }

    // Days between the 1904 epoch (1904-01-01) and the 1900 epoch that DateTime.FromOADate uses.
    private const double Date1904Offset = 1462;

    private static CellValue ParseCell(XElement cEl, IReadOnlyList<string> sharedStrings, HashSet<int> dateStyles, bool date1904)
    {
        var type = (string?)cEl.Attribute("t");

        if (type == "inlineStr")
        {
            var isEl = Elements(cEl, "is").FirstOrDefault();
            var text = isEl is null ? string.Empty : ExtractText(isEl);
            return CellValue.OfText(text);
        }

        var vEl = Elements(cEl, "v").FirstOrDefault();
        if (vEl is null) return CellValue.Empty;
        var raw = vEl.Value;

        switch (type)
        {
            case "s": // shared string: value is an index
                return int.TryParse(raw, out var si) && si >= 0 && si < sharedStrings.Count
                    ? CellValue.OfText(sharedStrings[si])
                    : CellValue.Empty;
            case "str": // formula string result
                return CellValue.OfText(raw);
            case "b": // boolean
                return CellValue.OfBoolean(raw == "1");
            case "e": // error
                return CellValue.OfText(raw);
            default: // number — possibly a date depending on the cell's style
                if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var num))
                {
                    var styleAttr = (string?)cEl.Attribute("s");
                    if (styleAttr != null && int.TryParse(styleAttr, out var styleIdx) && dateStyles.Contains(styleIdx))
                    {
                        var oaDate = date1904 ? num + Date1904Offset : num;
                        try { return CellValue.OfDate(DateTime.FromOADate(oaDate)); }
                        catch { /* out-of-range serial: fall through to number */ }
                    }
                    return CellValue.OfNumber(num);
                }
                return CellValue.OfText(raw);
        }
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        var result = new List<string>();
        var entry = GetEntry(archive, "xl/sharedStrings.xml");
        if (entry is null) return result;

        using var s = entry.Open();
        var doc = LoadPart(s);
        foreach (var si in Descendants(doc.Root, "si"))
            result.Add(ExtractText(si));
        return result;
    }

    /// <summary>
    /// Text of a shared-string or inline-string element (both use the CT_Rst content model): a
    /// direct &lt;t&gt; child, or the &lt;t&gt; of each &lt;r&gt; rich-text run, concatenated in
    /// order. Skips &lt;rPh&gt; phonetic-guide runs, which also carry a &lt;t&gt; but are not part
    /// of the string's text (issue #4).
    /// </summary>
    private static string ExtractText(XElement rstEl)
    {
        var texts = new List<string>();
        var directT = Elements(rstEl, "t").FirstOrDefault();
        if (directT != null) texts.Add(directT.Value);
        foreach (var r in Elements(rstEl, "r"))
        {
            var rt = Elements(r, "t").FirstOrDefault();
            if (rt != null) texts.Add(rt.Value);
        }
        return string.Concat(texts);
    }

    /// <summary>
    /// Returns the set of cell-style indices (positions in cellXfs) that use a date number format.
    /// </summary>
    private static HashSet<int> ReadDateStyles(ZipArchive archive)
    {
        var dateStyleIndices = new HashSet<int>();
        var entry = GetEntry(archive, "xl/styles.xml");
        if (entry is null) return dateStyleIndices;

        using var s = entry.Open();
        var doc = LoadPart(s);

        // Custom formats (id >= 164) whose format code looks like a date.
        var dateFormatIds = new HashSet<int>(BuiltinDateFormats);
        var numFmtsEl = Descendants(doc.Root, "numFmts").FirstOrDefault();
        if (numFmtsEl != null)
        {
            foreach (var fmt in Elements(numFmtsEl, "numFmt"))
            {
                var id = (int?)fmt.Attribute("numFmtId");
                var code = (string?)fmt.Attribute("formatCode");
                if (id is int fid && code != null && LooksLikeDate(code))
                    dateFormatIds.Add(fid);
            }
        }

        // cellXfs order == the style index a cell's "s" attribute points at.
        var cellXfs = Descendants(doc.Root, "cellXfs").FirstOrDefault();
        if (cellXfs != null)
        {
            var index = 0;
            foreach (var xf in Elements(cellXfs, "xf"))
            {
                var numFmtId = (int?)xf.Attribute("numFmtId") ?? 0;
                if (dateFormatIds.Contains(numFmtId))
                    dateStyleIndices.Add(index);
                index++;
            }
        }

        return dateStyleIndices;
    }

    /// <summary>
    /// True if the workbook uses the 1904 date system (`&lt;workbookPr date1904="1"/&gt;` in
    /// xl/workbook.xml). date1904 is xsd:boolean, so "1" and "true" both count.
    /// </summary>
    private static bool ReadDate1904(ZipArchive archive)
    {
        var entry = GetEntry(archive, "xl/workbook.xml");
        if (entry is null) return false;

        using var s = entry.Open();
        var doc = LoadPart(s);
        var attr = (string?)Descendants(doc.Root, "workbookPr").FirstOrDefault()?.Attribute("date1904");
        return attr == "1" || string.Equals(attr, "true", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeDate(string formatCode)
    {
        // Strip quoted literals and bracket sections, then look for date tokens. 'm' is ambiguous
        // (month vs minute) so we key off 'y' and 'd', which unambiguously indicate a date.
        var inBracket = false;
        var inQuote = false;
        foreach (var ch in formatCode)
        {
            if (ch == '"') { inQuote = !inQuote; continue; }
            if (inQuote) continue;
            if (ch == '[') { inBracket = true; continue; }
            if (ch == ']') { inBracket = false; continue; }
            if (inBracket) continue;
            if (ch is 'y' or 'Y' or 'd' or 'D') return true;
        }
        return false;
    }

    private static string ResolveFirstSheetPath(ZipArchive archive)
    {
        // Map the first <sheet> in workbook.xml to its part via the workbook relationships.
        var wb = GetEntry(archive, "xl/workbook.xml");
        var rels = GetEntry(archive, "xl/_rels/workbook.xml.rels");
        if (wb != null && rels != null)
        {
            try
            {
                using var wbStream = wb.Open();
                var wbDoc = LoadPart(wbStream);
                var firstSheet = Descendants(wbDoc.Root, "sheet").FirstOrDefault();
                var rid = firstSheet?.Attributes()
                    .FirstOrDefault(a => a.Name.LocalName == "id")?.Value; // r:id

                if (rid != null)
                {
                    using var relStream = rels.Open();
                    var relDoc = LoadPart(relStream);
                    var target = Descendants(relDoc.Root, "Relationship")
                        .FirstOrDefault(r => (string?)r.Attribute("Id") == rid)?
                        .Attribute("Target")?.Value;
                    if (!string.IsNullOrEmpty(target))
                        return target!.StartsWith("/") ? target!.TrimStart('/') : "xl/" + target;
                }
            }
            // Only fall back on a benign lookup failure (a missing rel, an unexpected shape).
            // A DOCTYPE is rejected by LoadPart and must not be swallowed into a silent fallback
            // that reads a different worksheet than the caller asked for.
            catch (Exception e) when (e is not XmlException) { /* fall through to convention */ }
        }

        // Fallback: first worksheet part by name.
        var sheet = archive.Entries
            .Where(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                        && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
        return sheet?.FullName ?? "xl/worksheets/sheet1.xml";
    }

    // ---- helpers ---------------------------------------------------------

    private static ZipArchiveEntry? GetEntry(ZipArchive archive, string path) =>
        archive.Entries.FirstOrDefault(e => string.Equals(e.FullName, path, StringComparison.OrdinalIgnoreCase));

    private static IEnumerable<XElement> Elements(XElement? parent, string localName) =>
        parent?.Elements().Where(e => e.Name.LocalName == localName) ?? Enumerable.Empty<XElement>();

    private static IEnumerable<XElement> Descendants(XElement? parent, string localName) =>
        parent?.Descendants().Where(e => e.Name.LocalName == localName) ?? Enumerable.Empty<XElement>();

    /// <summary>Zero-based row index from a row's "r" attribute (1-based in the file). -1 if absent
    /// or not a positive integer, in which case the row is appended at its current position. Throws
    /// if r is above Excel's maximum row, so a hostile value can't pad an unbounded grid.</summary>
    private static int RowIndex(string? rowRef)
    {
        if (!int.TryParse(rowRef, out var r) || r <= 0) return -1;
        if (r > MaxRowNumber)
            throw new InvalidDataException($"Row reference out of range: row {r} exceeds the maximum row {MaxRowNumber}.");
        return r - 1;
    }

    // Excel's last column is XFD, a zero-based index of 16383.
    private const int MaxColumnIndex = 16383;

    /// <summary>Zero-based column index from a cell reference like "AB12" (=> 27). -1 if absent. Throws
    /// <see cref="InvalidDataException"/>, naming the reference, if the column named is past Excel's
    /// maximum (XFD): that is not a column Excel could have written, and letting the accumulator run
    /// is how a crafted reference used to overflow to a negative number.</summary>
    public static int ColumnIndex(string? cellRef)
    {
        if (string.IsNullOrEmpty(cellRef)) return -1;
        var index = 0;
        var any = false;
        foreach (var ch in cellRef)
        {
            if (ch is >= 'A' and <= 'Z')
            {
                index = index * 26 + (ch - 'A' + 1);
                any = true;
            }
            else if (ch is >= 'a' and <= 'z')
            {
                index = index * 26 + (ch - 'a' + 1);
                any = true;
            }
            else break;

            // Checked every character, so the accumulator can never overflow before this runs.
            if (index > MaxColumnIndex + 1)
                throw new InvalidDataException($"Cell reference '{cellRef}' names a column past Excel's maximum (XFD).");
        }
        return any ? index - 1 : -1;
    }

    private static MemoryStream Buffer(Stream source)
    {
        var ms = new MemoryStream();
        source.CopyTo(ms);
        ms.Position = 0;
        return ms;
    }
}
