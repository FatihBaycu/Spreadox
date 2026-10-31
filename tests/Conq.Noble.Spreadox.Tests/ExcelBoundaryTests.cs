using System;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 8: Excel Limits and Boundary Tests (P0)
/// Verifies behavior at exact specification boundaries.
/// </summary>
public class ExcelBoundaryTests
{
    private readonly ITestOutputHelper _output;

    public ExcelBoundaryTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void ExactBoundary_RowAndColumnLimits()
    {
        // 1. Exact valid boundaries: 1,048,576 rows and 16,384 columns ("XFD1048576")
        CellAddress.Parse("XFD1048576", out int rMax, out int cMax);
        Assert.Equal(CellAddress.MaxRows, rMax);
        Assert.Equal(CellAddress.MaxColumns, cMax);
        Assert.Equal("XFD1048576", CellAddress.Format(CellAddress.MaxRows, CellAddress.MaxColumns));

        // 2. Exact valid minimum boundary ("A1")
        CellAddress.Parse("A1", out int rMin, out int cMin);
        Assert.Equal(1, rMin);
        Assert.Equal(1, cMin);
        Assert.Equal("A1", CellAddress.Format(1, 1));
    }

    [Theory]
    [InlineData("A0")]          // Row 0 invalid
    [InlineData("A1048577")]    // Row 1,048,577 exceeds max by 1
    [InlineData("XFE1")]        // Column 16,385 (XFE) exceeds max by 1
    [InlineData("@1")]          // Invalid column char
    public void OutOfBoundsCoordinates_ThrowArgumentException_WithDescriptiveMessage(string address)
    {
        var ex = Assert.Throws<ArgumentException>(() => CellAddress.Parse(address, out _, out _));
        Assert.NotEmpty(ex.Message);
        _output.WriteLine($"Address '{address}' correctly rejected: {ex.Message}");
    }

    [Fact]
    public void WorksheetName_Exact31Characters_Succeeds_And32CharactersFails()
    {
        using var wb = ExcelWorkbook.Create();

        // Exactly 31 characters
        string name31 = new string('A', 31);
        var sheet31 = wb.AddWorksheet(name31);
        Assert.Equal(name31, sheet31.Name);

        // 32 characters
        string name32 = new string('B', 32);
        var ex = Assert.Throws<ArgumentException>(() => wb.AddWorksheet(name32));
        Assert.Contains("31", ex.Message);
    }

    [Theory]
    [InlineData("Sheet/1")]
    [InlineData("Sheet\\1")]
    [InlineData("Sheet?1")]
    [InlineData("Sheet*1")]
    [InlineData("Sheet[1]")]
    [InlineData("Sheet:1")]
    [InlineData("'Sheet")]
    [InlineData("Sheet'")]
    public void WorksheetName_ForbiddenCharacters_ThrowArgumentException(string name)
    {
        using var wb = ExcelWorkbook.Create();
        Assert.Throws<ArgumentException>(() => wb.AddWorksheet(name));
    }
}
