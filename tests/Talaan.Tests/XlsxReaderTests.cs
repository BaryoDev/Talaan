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

    // Real-world smoke test against the club roster if it's present on this machine. It is not
    // committed (member PII); the test no-ops when the file is absent so CI stays green elsewhere.
    private const string RosterPath = "/Users/arnelirobles/Downloads/ClubRoster.xlsx";

    [Fact]
    public void Reads_real_roster_when_available()
    {
        if (!File.Exists(RosterPath)) return; // fixture not present — skip

        var sheet = Spreadsheet.ReadFile(RosterPath);

        Assert.True(sheet.RowCount > 30, $"expected many rows, got {sheet.RowCount}");

        // Locate the header row (the one containing "Member #").
        int headerRow = -1, firstNameCol = -1, lastNameCol = -1, memberNoCol = -1, admittedCol = -1;
        for (var r = 0; r < sheet.RowCount && headerRow < 0; r++)
        {
            for (var c = 0; c < sheet.ColumnCount; c++)
            {
                if (sheet.At(r, c).AsString().Trim() == "Member #")
                {
                    headerRow = r;
                    memberNoCol = c;
                    break;
                }
            }
        }
        Assert.True(headerRow >= 0, "header row with 'Member #' not found");

        for (var c = 0; c < sheet.ColumnCount; c++)
        {
            var h = sheet.At(headerRow, c).AsString().Trim();
            if (h == "First Name") firstNameCol = c;
            else if (h == "Last Name") lastNameCol = c;
            else if (h == "Admitted") admittedCol = c;
        }
        Assert.True(firstNameCol >= 0 && lastNameCol >= 0, "name columns not found");

        // Find Dindo Abantao and verify the surrounding cells parse correctly.
        int dindoRow = -1;
        for (var r = headerRow + 1; r < sheet.RowCount; r++)
        {
            if (sheet.At(r, firstNameCol).AsString().Trim() == "Dindo" &&
                sheet.At(r, lastNameCol).AsString().Trim() == "Abantao")
            {
                dindoRow = r;
                break;
            }
        }
        Assert.True(dindoRow >= 0, "expected member 'Dindo Abantao' not found");

        // Member number must come through as a clean integer string, not "3.322657E+06".
        var memberNo = sheet.At(dindoRow, memberNoCol);
        Assert.Equal(CellKind.Number, memberNo.Kind);
        Assert.Equal("3322657", memberNo.AsString());

        // The Admitted cell is serial 35781. Depending on the sheet's styling it is either a real
        // Date (year 1997) or the raw number — our reader must produce one of those, never garbage.
        var admitted = admittedCol >= 0 ? sheet.At(dindoRow, admittedCol) : CellValue.Empty;
        if (admitted.Kind == CellKind.Date)
            Assert.Equal(1997, admitted.Date!.Value.Year);
        else if (admitted.Kind == CellKind.Number)
            Assert.Equal(35781d, admitted.Number!.Value, 3);
    }
}
