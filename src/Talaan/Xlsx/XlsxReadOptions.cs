namespace Talaan.Xlsx;

/// <summary>
/// Optional limits for reading an .xlsx file. Every limit is off by default, so a caller that does
/// not pass options sees the same behaviour as before.
/// </summary>
public sealed class XlsxReadOptions
{
    /// <summary>
    /// The most cells a worksheet may produce, counting the empty cells Talaan pads in so every row
    /// lines up with its widest cell. <c>null</c> (the default) means no limit.
    /// <para>
    /// A sheet that holds one cell at column XFD in each row is a few KB on disk but reads as 16384
    /// cells per row, about 1.75 MB of memory per row. A byte limit does not catch that, a cell
    /// limit does. The check runs before the padding is allocated, and exceeding it throws
    /// <see cref="InvalidDataException"/>.
    /// </para>
    /// </summary>
    public long? MaxCells { get; init; }
}
