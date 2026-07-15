using System.Globalization;

namespace Talaan;

/// <summary>The intrinsic type a parsed cell resolved to.</summary>
public enum CellKind
{
    Empty,
    Text,
    Number,
    Date,
    Boolean
}

/// <summary>
/// A single parsed cell. Only the field matching <see cref="Kind"/> is populated. Value types keep
/// the reader allocation-light and let callers branch on <see cref="Kind"/> rather than re-parsing text.
/// </summary>
public readonly record struct CellValue(CellKind Kind, string? Text, double? Number, DateTime? Date, bool? Boolean)
{
    public static readonly CellValue Empty = new(CellKind.Empty, null, null, null, null);

    public static CellValue OfText(string value) =>
        string.IsNullOrEmpty(value) ? Empty : new(CellKind.Text, value, null, null, null);

    public static CellValue OfNumber(double value) => new(CellKind.Number, null, value, null, null);
    public static CellValue OfDate(DateTime value) => new(CellKind.Date, null, null, value, null);
    public static CellValue OfBoolean(bool value) => new(CellKind.Boolean, null, null, null, value);

    /// <summary>True for an empty cell or a text cell that is blank/whitespace.</summary>
    public bool IsBlank => Kind == CellKind.Empty || (Kind == CellKind.Text && string.IsNullOrWhiteSpace(Text));

    /// <summary>Best-effort display string. Numbers use invariant culture; dates use ISO-8601 date.</summary>
    public string AsString() => Kind switch
    {
        CellKind.Text => Text ?? string.Empty,
        CellKind.Number => Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
        CellKind.Date => Date?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? string.Empty,
        CellKind.Boolean => Boolean == true ? "TRUE" : "FALSE",
        _ => string.Empty
    };

    public override string ToString() => AsString();
}
