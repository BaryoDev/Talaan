using System.IO.Compression;
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

    public static SheetData Read(Stream stream)
    {
        // ZipArchive needs a seekable stream; buffer if necessary.
        Stream seekable = stream.CanSeek ? stream : Buffer(stream);
        using var archive = new ZipArchive(seekable, ZipArchiveMode.Read, leaveOpen: true);

        var sharedStrings = ReadSharedStrings(archive);
        var dateStyles = ReadDateStyles(archive);
        var sheetPath = ResolveFirstSheetPath(archive);

        var entry = GetEntry(archive, sheetPath)
            ?? throw new InvalidDataException($"Worksheet part '{sheetPath}' not found in workbook.");

        using var sheetStream = entry.Open();
        var doc = XDocument.Load(sheetStream);

        var rows = new List<IReadOnlyList<CellValue>>();
        foreach (var rowEl in Descendants(doc.Root, "row"))
        {
            var cells = new List<CellValue>();
            foreach (var cEl in Elements(rowEl, "c"))
            {
                var colIndex = ColumnIndex((string?)cEl.Attribute("r"));
                // Pad any skipped columns with empties so alignment is preserved.
                while (colIndex >= 0 && cells.Count < colIndex) cells.Add(CellValue.Empty);
                cells.Add(ParseCell(cEl, sharedStrings, dateStyles));
            }
            rows.Add(cells);
        }

        return new SheetData(rows, name: null);
    }

    private static CellValue ParseCell(XElement cEl, IReadOnlyList<string> sharedStrings, HashSet<int> dateStyles)
    {
        var type = (string?)cEl.Attribute("t");

        if (type == "inlineStr")
        {
            var isEl = Elements(cEl, "is").FirstOrDefault();
            var text = isEl is null ? string.Empty : string.Concat(Descendants(isEl, "t").Select(t => t.Value));
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
                        try { return CellValue.OfDate(DateTime.FromOADate(num)); }
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
        var doc = XDocument.Load(s);
        foreach (var si in Descendants(doc.Root, "si"))
        {
            // <si> may hold a single <t> or several rich-text <r><t> runs; concatenate all <t>.
            result.Add(string.Concat(Descendants(si, "t").Select(t => t.Value)));
        }
        return result;
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
        var doc = XDocument.Load(s);

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
                var wbDoc = XDocument.Load(wbStream);
                var firstSheet = Descendants(wbDoc.Root, "sheet").FirstOrDefault();
                var rid = firstSheet?.Attributes()
                    .FirstOrDefault(a => a.Name.LocalName == "id")?.Value; // r:id

                if (rid != null)
                {
                    using var relStream = rels.Open();
                    var relDoc = XDocument.Load(relStream);
                    var target = Descendants(relDoc.Root, "Relationship")
                        .FirstOrDefault(r => (string?)r.Attribute("Id") == rid)?
                        .Attribute("Target")?.Value;
                    if (!string.IsNullOrEmpty(target))
                        return target!.StartsWith("/") ? target!.TrimStart('/') : "xl/" + target;
                }
            }
            catch { /* fall through to convention */ }
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

    /// <summary>Zero-based column index from a cell reference like "AB12" (=> 27). -1 if absent.</summary>
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
