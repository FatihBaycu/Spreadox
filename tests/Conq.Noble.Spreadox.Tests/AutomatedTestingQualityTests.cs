using System;
using System.IO;
using Conq.Noble.Spreadox;
using Xunit;

namespace Conq.Noble.Spreadox.Tests;

public class AutomatedTestingQualityTests
{
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void AddWorksheet_NullOrWhitespaceName_ThrowsArgumentException(string? invalidName)
    {
        using var wb = ExcelWorkbook.Create();
        Assert.ThrowsAny<ArgumentException>(() => wb.AddWorksheet(invalidName!));
    }

    [Theory]
    [InlineData("Sheet[1]")]
    [InlineData("Sheet?1")]
    [InlineData("Sheet/1")]
    [InlineData("Sheet\\1")]
    [InlineData("Sheet*1")]
    [InlineData("Sheet:1")]
    [InlineData("'Sheet")]
    [InlineData("Sheet'")]
    public void AddWorksheet_InvalidCharacters_ThrowsArgumentException(string invalidName)
    {
        using var wb = ExcelWorkbook.Create();
        Assert.ThrowsAny<ArgumentException>(() => wb.AddWorksheet(invalidName));
    }

    [Fact]
    public void AddWorksheet_Exceeding31Characters_ThrowsArgumentException()
    {
        using var wb = ExcelWorkbook.Create();
        string longName = new string('A', 32);
        Assert.Throws<ArgumentException>(() => wb.AddWorksheet(longName));
    }

    [Fact]
    public void AddWorksheet_DuplicateName_ThrowsInvalidOperationException()
    {
        using var wb = ExcelWorkbook.Create();
        wb.AddWorksheet("TestSheet");
        Assert.Throws<InvalidOperationException>(() => wb.AddWorksheet("testsheet")); // Case-insensitive duplicate
    }

    [Fact]
    public void RemoveLastWorksheet_ThrowsInvalidOperationException()
    {
        using var wb = ExcelWorkbook.Create();
        Assert.Single(wb.Worksheets);
        Assert.Throws<InvalidOperationException>(() => wb.RemoveWorksheet(wb.Worksheets[0]));
    }

    [Fact]
    public void CellAddress_BoundaryCoordinates()
    {
        // Min valid: Row 1, Column 1 ("A1")
        CellAddress.Parse("A1", out int rowMin, out int colMin);
        Assert.Equal(1, rowMin);
        Assert.Equal(1, colMin);

        // Max valid: Row 1,048,576, Column 16,384 ("XFD1048576")
        CellAddress.Parse("XFD1048576", out int rowMax, out int colMax);
        Assert.Equal(1048576, rowMax);
        Assert.Equal(16384, colMax);

        // Boundary formatting check
        Assert.Equal("XFD1048576", CellAddress.Format(1048576, 16384));
    }

    [Theory]
    [InlineData("A0")]
    [InlineData("A1048577")] // Row beyond max
    [InlineData("XFE1")]     // Column beyond XFD
    [InlineData("")]
    [InlineData("Invalid")]
    public void CellAddress_InvalidInputs_ThrowsArgumentException(string invalidAddress)
    {
        Assert.Throws<ArgumentException>(() => CellAddress.Parse(invalidAddress, out _, out _));
    }

    [Fact]
    public void ExcelStreamingWorkbook_AddingDuplicateWorksheet_ThrowsInvalidOperationException()
    {
        using var ms = new MemoryStream();
        using var sw = ExcelStreamingWorkbook.Create(ms);
        sw.AddWorksheet("Data");
        Assert.Throws<InvalidOperationException>(() => sw.AddWorksheet("data"));
    }
}

