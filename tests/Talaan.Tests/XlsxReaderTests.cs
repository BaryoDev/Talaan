using Talaan;
using Talaan.Xlsx;
using Xunit;

namespace Talaan.Tests;

public class XlsxReaderTests
{
    [Theory]
    [InlineData("A1", 0)]
    [InlineData("B1", 1)]
    [InlineData("Z9", 25)]
    [InlineData("AA1", 26)]
    [InlineData("AB12", 27)]
    public void Column_index_from_reference(string cellRef, int expected)
        => Assert.Equal(expected, XlsxReader.ColumnIndex(cellRef));

    // NOTE: the xlsx reader has no fixture-backed coverage yet. The previous smoke test read a
    // roster from a hard-coded local path, so it silently no-opped everywhere else and asserted
    // against real personal data. Replace it with small generated .xlsx fixtures covering shared
    // and inline strings, row and column gaps, date styles, the 1904 date system, booleans and
    // formula results.
}
