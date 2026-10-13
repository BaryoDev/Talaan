using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Talaan.Xlsx;

/// <summary>
/// Writes a minimal .xlsx package using only <see cref="ZipArchive"/> and <see cref="XmlWriter"/>.
/// Text is written as inline strings (no shared strings table). Dates are written as serial numbers
/// with one of two styles: <c>yyyy-mm-dd</c> when the value has no time, <c>yyyy-mm-dd hh:mm:ss</c>
/// when it does.
/// </summary>
internal static class XlsxWriter
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    private const int StyleDate = 1;
    private const int StyleDateTime = 2;

    // Excel treats 1900 as a leap year, so serials before 1900-03-01 differ by a day from .NET's
    // OLE Automation dates. Refuse them rather than write a date that shows up one day off.
    private static readonly DateTime EarliestDate = new(1900, 3, 1);

    private const int MaxRows = 1_048_576;
    private const int MaxColumns = 16_384;
    private const int MaxTextLength = 32_767;

    private static readonly char[] ForbiddenNameChars = { '[', ']', ':', '*', '?', '/', '\\' };

    private static readonly XmlWriterSettings WriterSettings = new()
    {
        Encoding = new UTF8Encoding(false),
        // Turn CR into &#xD; so it survives a read. The default would turn a lone CR into CRLF.
        NewLineHandling = NewLineHandling.Entitize,
        CloseOutput = true,
    };

    public static void Write(Stream stream, IEnumerable<SheetData> sheets)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(sheets);

        var list = sheets.ToList();
        if (list.Count == 0)
            throw new ArgumentException("At least one sheet is required.", nameof(sheets));
        if (list.Any(s => s is null))
            throw new ArgumentException("The sheet list contains a null sheet.", nameof(sheets));

        var names = ResolveNames(list);
        Validate(list, names);

        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        WritePart(zip, "[Content_Types].xml", w =>
        {
            w.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");
            Default(w, "rels", "application/vnd.openxmlformats-package.relationships+xml");
            Default(w, "xml", "application/xml");
            Override(w, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            for (var i = 1; i <= list.Count; i++)
                Override(w, $"/xl/worksheets/sheet{i}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            Override(w, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            w.WriteEndElement();
        });

        WritePart(zip, "_rels/.rels", w =>
        {
            w.WriteStartElement("Relationships", PkgRelNs);
            Relationship(w, "rId1", RelNs + "/officeDocument", "xl/workbook.xml");
            w.WriteEndElement();
        });

        WritePart(zip, "xl/workbook.xml", w =>
        {
            w.WriteStartElement("workbook", MainNs);
            w.WriteAttributeString("xmlns", "r", null, RelNs);
            w.WriteStartElement("sheets");
            for (var i = 0; i < names.Count; i++)
            {
                w.WriteStartElement("sheet");
                w.WriteAttributeString("name", names[i]);
                w.WriteAttributeString("sheetId", (i + 1).ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("r", "id", RelNs, $"rId{i + 1}");
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndElement();
        });

        WritePart(zip, "xl/_rels/workbook.xml.rels", w =>
        {
            w.WriteStartElement("Relationships", PkgRelNs);
            for (var i = 1; i <= list.Count; i++)
                Relationship(w, $"rId{i}", RelNs + "/worksheet", $"worksheets/sheet{i}.xml");
            Relationship(w, $"rId{list.Count + 1}", RelNs + "/styles", "styles.xml");
            w.WriteEndElement();
        });

        WritePart(zip, "xl/styles.xml", WriteStyles);

        for (var i = 0; i < list.Count; i++)
        {
            var sheet = list[i];
            WritePart(zip, $"xl/worksheets/sheet{i + 1}.xml", w => WriteSheet(w, sheet));
        }
    }

    /// <summary>
    /// Checks every sheet and cell before anything is written, so a rejection leaves the caller's
    /// stream untouched.
    /// </summary>
    private static void Validate(List<SheetData> sheets, List<string> names)
    {
        for (var i = 0; i < sheets.Count; i++)
        {
            var sheet = sheets[i];
            var name = names[i];
            if (sheet.Rows.Count > MaxRows)
                throw new ArgumentException($"Sheet '{name}' has {sheet.Rows.Count} rows. A sheet can have at most {MaxRows}.", "sheets");

            for (var r = 0; r < sheet.Rows.Count; r++)
            {
                var row = sheet.Rows[r];
                if (row.Count > MaxColumns)
                    throw new ArgumentException($"Sheet '{name}' row {r + 1} has {row.Count} cells. A sheet can have at most {MaxColumns} columns (XFD).", "sheets");

                for (var c = 0; c < row.Count; c++)
                {
                    var cell = row[c];
                    var where = $"Sheet '{name}' cell {ColumnName(c)}{r + 1}";
                    switch (cell.Kind)
                    {
                        case CellKind.Text:
                            var len = StripIllegalXmlChars(cell.Text ?? string.Empty).Length;
                            if (len > MaxTextLength)
                                throw new ArgumentException($"{where} holds {len} characters. A cell can hold at most {MaxTextLength}.", "sheets");
                            break;
                        case CellKind.Number:
                            var n = cell.Number ?? 0d;
                            if (double.IsNaN(n) || double.IsInfinity(n))
                                throw new ArgumentException($"{where} holds {n}, which a workbook cannot store.", "sheets");
                            break;
                        case CellKind.Date:
                            var d = cell.Date ?? default;
                            if (d < EarliestDate)
                                throw new ArgumentException($"{where} holds {d:yyyy-MM-dd}. Dates before 1900-03-01 are not supported.", "sheets");
                            break;
                    }
                }
            }
        }
    }

    private static List<string> ResolveNames(List<SheetData> sheets)
    {
        var names = new List<string>(sheets.Count);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < sheets.Count; i++)
        {
            var name = StripIllegalXmlChars(sheets[i].Name ?? string.Empty);
            if (string.IsNullOrWhiteSpace(name))
                name = $"Sheet{i + 1}";

            if (name.Length > 31)
                throw new ArgumentException(
                    $"Sheet name '{name}' is {name.Length} characters. Excel allows at most 31.", nameof(sheets));
            var ctl = name.IndexOfAny(new[] { '\t', '\r', '\n' });
            if (ctl >= 0)
                throw new ArgumentException(
                    $"Sheet name '{name}' contains a tab, CR or LF, which is not allowed.", nameof(sheets));
            if (name.Equals("History", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    "Sheet name 'History' is reserved by Excel (compared without regard to case).", nameof(sheets));
            var bad = name.IndexOfAny(ForbiddenNameChars);
            if (bad >= 0)
                throw new ArgumentException(
                    $"Sheet name '{name}' contains '{name[bad]}'. These characters are not allowed: [ ] : * ? / \\", nameof(sheets));
            if (name[0] == '\'' || name[^1] == '\'')
                throw new ArgumentException(
                    $"Sheet name '{name}' cannot start or end with an apostrophe.", nameof(sheets));
            if (!seen.Add(name))
                throw new ArgumentException(
                    $"Duplicate sheet name '{name}'. Sheet names are compared without regard to case.", nameof(sheets));

            names.Add(name);
        }
        return names;
    }

    private static void WriteSheet(XmlWriter w, SheetData sheet)
    {
        w.WriteStartElement("worksheet", MainNs);
        w.WriteStartElement("sheetData");
        for (var r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            var rowOpen = false;
            for (var c = 0; c < row.Count; c++)
            {
                var cell = row[c];
                if (cell.Kind == CellKind.Empty) continue;
                if (cell.Kind == CellKind.Text && string.IsNullOrEmpty(cell.Text)) continue;

                if (!rowOpen)
                {
                    w.WriteStartElement("row");
                    w.WriteAttributeString("r", (r + 1).ToString(CultureInfo.InvariantCulture));
                    rowOpen = true;
                }
                WriteCell(w, cell, ColumnName(c) + (r + 1).ToString(CultureInfo.InvariantCulture));
            }
            if (rowOpen) w.WriteEndElement();
        }
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void WriteCell(XmlWriter w, CellValue cell, string reference)
    {
        switch (cell.Kind)
        {
            case CellKind.Text:
            {
                var text = StripIllegalXmlChars(cell.Text!);
                if (text.Length == 0) return;
                w.WriteStartElement("c");
                w.WriteAttributeString("r", reference);
                w.WriteAttributeString("t", "inlineStr");
                w.WriteStartElement("is");
                w.WriteStartElement("t");
                w.WriteAttributeString("xml", "space", null, "preserve");
                w.WriteString(text);
                w.WriteEndElement();
                w.WriteEndElement();
                w.WriteEndElement();
                break;
            }
            case CellKind.Number:
            {
                var n = cell.Number ?? 0d;
                w.WriteStartElement("c");
                w.WriteAttributeString("r", reference);
                w.WriteElementString("v", n.ToString("R", CultureInfo.InvariantCulture));
                w.WriteEndElement();
                break;
            }
            case CellKind.Boolean:
                w.WriteStartElement("c");
                w.WriteAttributeString("r", reference);
                w.WriteAttributeString("t", "b");
                w.WriteElementString("v", cell.Boolean == true ? "1" : "0");
                w.WriteEndElement();
                break;
            case CellKind.Date:
            {
                var d = cell.Date ?? default;
                w.WriteStartElement("c");
                w.WriteAttributeString("r", reference);
                w.WriteAttributeString("s", (d.TimeOfDay == TimeSpan.Zero ? StyleDate : StyleDateTime).ToString(CultureInfo.InvariantCulture));
                w.WriteElementString("v", d.ToOADate().ToString("R", CultureInfo.InvariantCulture));
                w.WriteEndElement();
                break;
            }
        }
    }

    private static void WriteStyles(XmlWriter w)
    {
        w.WriteStartElement("styleSheet", MainNs);

        w.WriteStartElement("numFmts");
        w.WriteAttributeString("count", "2");
        NumFmt(w, 164, "yyyy\\-mm\\-dd");
        NumFmt(w, 165, "yyyy\\-mm\\-dd\\ hh:mm:ss");
        w.WriteEndElement();

        w.WriteStartElement("fonts");
        w.WriteAttributeString("count", "1");
        w.WriteStartElement("font");
        w.WriteStartElement("sz"); w.WriteAttributeString("val", "11"); w.WriteEndElement();
        w.WriteStartElement("name"); w.WriteAttributeString("val", "Calibri"); w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteStartElement("fills");
        w.WriteAttributeString("count", "2");
        Fill(w, "none");
        Fill(w, "gray125");
        w.WriteEndElement();

        w.WriteStartElement("borders");
        w.WriteAttributeString("count", "1");
        w.WriteStartElement("border");
        foreach (var side in new[] { "left", "right", "top", "bottom", "diagonal" })
            w.WriteElementString(side, null);
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteStartElement("cellStyleXfs");
        w.WriteAttributeString("count", "1");
        Xf(w, 0, withApply: false);
        w.WriteEndElement();

        w.WriteStartElement("cellXfs");
        w.WriteAttributeString("count", "3");
        Xf(w, 0, withApply: false, xfId: true);
        Xf(w, 164, withApply: true, xfId: true);
        Xf(w, 165, withApply: true, xfId: true);
        w.WriteEndElement();

        w.WriteStartElement("cellStyles");
        w.WriteAttributeString("count", "1");
        w.WriteStartElement("cellStyle");
        w.WriteAttributeString("name", "Normal");
        w.WriteAttributeString("xfId", "0");
        w.WriteAttributeString("builtinId", "0");
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteEndElement();
    }

    private static void NumFmt(XmlWriter w, int id, string code)
    {
        w.WriteStartElement("numFmt");
        w.WriteAttributeString("numFmtId", id.ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("formatCode", code);
        w.WriteEndElement();
    }

    private static void Fill(XmlWriter w, string pattern)
    {
        w.WriteStartElement("fill");
        w.WriteStartElement("patternFill");
        w.WriteAttributeString("patternType", pattern);
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void Xf(XmlWriter w, int numFmtId, bool withApply, bool xfId = false)
    {
        w.WriteStartElement("xf");
        w.WriteAttributeString("numFmtId", numFmtId.ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("fontId", "0");
        w.WriteAttributeString("fillId", "0");
        w.WriteAttributeString("borderId", "0");
        if (xfId) w.WriteAttributeString("xfId", "0");
        if (withApply) w.WriteAttributeString("applyNumberFormat", "1");
        w.WriteEndElement();
    }

    private static void Default(XmlWriter w, string ext, string type)
    {
        w.WriteStartElement("Default");
        w.WriteAttributeString("Extension", ext);
        w.WriteAttributeString("ContentType", type);
        w.WriteEndElement();
    }

    private static void Override(XmlWriter w, string part, string type)
    {
        w.WriteStartElement("Override");
        w.WriteAttributeString("PartName", part);
        w.WriteAttributeString("ContentType", type);
        w.WriteEndElement();
    }

    private static void Relationship(XmlWriter w, string id, string type, string target)
    {
        w.WriteStartElement("Relationship");
        w.WriteAttributeString("Id", id);
        w.WriteAttributeString("Type", type);
        w.WriteAttributeString("Target", target);
        w.WriteEndElement();
    }

    private static void WritePart(ZipArchive zip, string path, Action<XmlWriter> body)
    {
        var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
        using var xml = XmlWriter.Create(entry.Open(), WriterSettings);
        xml.WriteStartDocument(standalone: true);
        body(xml);
        xml.WriteEndDocument();
    }

    /// <summary>Zero-based column index to letters: 0 is A, 25 is Z, 26 is AA.</summary>
    internal static string ColumnName(int index)
    {
        var sb = new StringBuilder();
        for (var n = index + 1; n > 0; n = (n - 1) / 26)
            sb.Insert(0, (char)('A' + (n - 1) % 26));
        return sb.ToString();
    }

    /// <summary>Drops characters XML 1.0 does not allow, including lone surrogates.</summary>
    internal static string StripIllegalXmlChars(string value)
    {
        StringBuilder? sb = null;
        for (var i = 0; i < value.Length; i++)
        {
            var ch = value[i];
            var keep = false;
            var pair = false;
            if (ch is '\t' or '\n' or '\r' || (ch >= 0x20 && ch <= 0xD7FF) || (ch >= 0xE000 && ch <= 0xFFFD))
                keep = true;
            else if (char.IsHighSurrogate(ch) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                keep = pair = true;

            if (keep)
            {
                sb?.Append(ch);
                if (pair) { i++; sb?.Append(value[i]); }
            }
            else
            {
                sb ??= new StringBuilder(value.Length).Append(value, 0, i);
            }
        }
        return sb?.ToString() ?? value;
    }
}
