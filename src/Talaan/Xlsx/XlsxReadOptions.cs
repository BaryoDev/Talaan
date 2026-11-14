namespace Talaan.Xlsx;

/// <summary>
/// Optional limits for reading an .xlsx file. Every limit is off by default, so a caller that does
/// not pass options sees the same behaviour as before.
/// </summary>
public sealed class XlsxReadOptions
{
    /// <summary>
    /// The most cells the returned grid may hold, counting the empty cells Talaan pads in so every
    /// row lines up with its last cell. <c>null</c> (the default) means no limit; zero or a negative
    /// value rejects the first cell. Exceeding it throws <see cref="InvalidDataException"/>.
    /// <para>
    /// A sheet with one cell at column XFD in each row is a few KB on disk but reads as 16384 cells
    /// per row (about 0.9 MB retained per row, since each cell is 56 bytes). The check runs before a
    /// row's padding is allocated. It bounds the output grid only: it does not bound padding for
    /// skipped rows (a single <c>&lt;row r="1048576"&gt;</c> still adds about a million empty rows) or
    /// the memory used to parse the XML parts.
    /// </para>
    /// </summary>
    public long? MaxCells { get; init; }
}
