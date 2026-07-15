namespace Talaan;

/// <summary>
/// A parsed sheet as a raw rectangular grid of cells. No header interpretation is applied — the
/// caller decides which row(s) are headers and which to skip. Rows may have differing lengths as
/// stored; <see cref="ColumnCount"/> reports the widest.
/// </summary>
public sealed class SheetData
{
    public string? Name { get; init; }

    public IReadOnlyList<IReadOnlyList<CellValue>> Rows { get; }

    public SheetData(IReadOnlyList<IReadOnlyList<CellValue>> rows, string? name = null)
    {
        Rows = rows;
        Name = name;
    }

    public int RowCount => Rows.Count;

    public int ColumnCount => Rows.Count == 0 ? 0 : Rows.Max(r => r.Count);

    /// <summary>Cell at (row, col) or <see cref="CellValue.Empty"/> if out of range.</summary>
    public CellValue At(int row, int col)
    {
        if (row < 0 || row >= Rows.Count) return CellValue.Empty;
        var r = Rows[row];
        return col < 0 || col >= r.Count ? CellValue.Empty : r[col];
    }
}
