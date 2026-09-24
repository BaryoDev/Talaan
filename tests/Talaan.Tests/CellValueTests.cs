using System;
using Talaan;
using Xunit;

namespace Talaan.Tests;

public class CellValueTests
{
    [Fact]
    public void Date_only_renders_as_yyyy_MM_dd()
    {
        var cell = CellValue.OfDate(new DateTime(2026, 9, 1));
        Assert.Equal("2026-09-01", cell.AsString());
    }

    [Fact]
    public void Date_with_time_renders_as_iso8601()
    {
        var cell = CellValue.OfDate(new DateTime(2026, 9, 1, 13, 45, 0));
        Assert.Equal("2026-09-01T13:45:00", cell.AsString());
    }

    [Fact]
    public void Time_only_renders_as_the_OLE_epoch_datetime()
    {
        // Serial 0.5 under a time-only format such as numFmtId 45 (mm:ss) decodes via
        // DateTime.FromOADate to noon on the OLE epoch, 1899-12-30. Per issue #5, AsString()
        // keeps the full datetime rather than trying to detect "time-only" (CellValue has no
        // numFmtId to do that with), so the string is verbose but not wrong.
        var cell = CellValue.OfDate(new DateTime(1899, 12, 30, 12, 0, 0));
        Assert.Equal("1899-12-30T12:00:00", cell.AsString());
    }

    [Fact]
    public void Midnight_exactly_has_no_time_component()
    {
        var cell = CellValue.OfDate(new DateTime(2026, 9, 1, 0, 0, 0));
        Assert.Equal("2026-09-01", cell.AsString());
    }

    [Fact]
    public void Number_renders_invariant()
    {
        var cell = CellValue.OfNumber(1234.5);
        Assert.Equal("1234.5", cell.AsString());
    }

    [Fact]
    public void Text_renders_as_is()
    {
        var cell = CellValue.OfText("hello");
        Assert.Equal("hello", cell.AsString());
    }

    [Fact]
    public void Boolean_renders_as_TRUE_or_FALSE()
    {
        Assert.Equal("TRUE", CellValue.OfBoolean(true).AsString());
        Assert.Equal("FALSE", CellValue.OfBoolean(false).AsString());
    }
}
