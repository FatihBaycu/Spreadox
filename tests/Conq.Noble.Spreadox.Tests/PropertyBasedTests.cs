using System;
using System.IO;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 2: Property-Based Testing (P0)
/// Automated invariant validation across thousands of randomized inputs.
/// </summary>
public class PropertyBasedTests
{
    private readonly ITestOutputHelper _output;

    public PropertyBasedTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void CellAddress_FormatAndParse_RoundTripInvariant_AcrossRandomCoordinates()
    {
        // Deterministic pseudo-random seed for reproducibility
        var rand = new Random(424242);
        const int iterations = 10_000;

        for (int i = 0; i < iterations; i++)
        {
            // Pick random valid rows (1..1,048,576) and columns (1..16,384)
            int row = rand.Next(1, CellAddress.MaxRows + 1);
            int col = rand.Next(1, CellAddress.MaxColumns + 1);

            string address = CellAddress.Format(row, col);
            CellAddress.Parse(address, out int parsedRow, out int parsedCol);

            Assert.Equal(row, parsedRow);
            Assert.Equal(col, parsedCol);
        }

        _output.WriteLine($"Successfully verified CellAddress round-trip invariant across {iterations:N0} randomized coordinates.");
    }

    [Fact]
    public void ExcelRange_ParseAndFormat_RoundTripInvariant()
    {
        var rand = new Random(1337);
        const int iterations = 2_000;

        for (int i = 0; i < iterations; i++)
        {
            int r1 = rand.Next(1, 10_000);
            int r2 = rand.Next(r1, 10_000);
            int c1 = rand.Next(1, 1_000);
            int c2 = rand.Next(c1, 1_000);

            string a1Expected = $"{CellAddress.Format(r1, c1)}:{CellAddress.Format(r2, c2)}";
            var range = ExcelRange.Parse(a1Expected);

            Assert.Equal(r1, range.StartRow);
            Assert.Equal(c1, range.StartColumn);
            Assert.Equal(r2, range.EndRow);
            Assert.Equal(c2, range.EndColumn);
            Assert.Equal(a1Expected, range.A1);
        }

        _output.WriteLine($"Successfully verified ExcelRange round-trip invariant across {iterations:N0} randomized 2D ranges.");
    }

    [Fact]
    public void ExcelColor_HexRoundTripInvariant()
    {
        var rand = new Random(999);
        const int iterations = 2_000;

        for (int i = 0; i < iterations; i++)
        {
            byte r = (byte)rand.Next(0, 256);
            byte g = (byte)rand.Next(0, 256);
            byte b = (byte)rand.Next(0, 256);

            var color = ExcelColor.FromRgb(r, g, b);
            string hex = color.ToRgbHex();
            var parsed = ExcelColor.FromHex(hex);

            Assert.Equal(r, parsed.R);
            Assert.Equal(g, parsed.G);
            Assert.Equal(b, parsed.B);
        }

        _output.WriteLine($"Successfully verified ExcelColor round-trip invariant across {iterations:N0} randomized RGB colors.");
    }

    [Fact]
    public void ExcelCell_NumericData_FidelityInvariant()
    {
        var rand = new Random(777);
        const int iterations = 1_000;

        using var wb = ExcelWorkbook.Create();
        var sheet = wb.Worksheets[0];

        for (int i = 1; i <= iterations; i++)
        {
            double val = (rand.NextDouble() - 0.5) * 1_000_000_000.0;
            sheet[i, 1].SetValue(val);
            Assert.Equal(val, sheet[i, 1].GetDouble());
        }

        _output.WriteLine($"Successfully verified numeric precision invariant across {iterations:N0} randomized cells.");
    }
}
