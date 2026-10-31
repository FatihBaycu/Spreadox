using System;
using System.IO;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 7: Decimal & Financial Precision Tests (P0)
/// Verifies financial calculations, decimals, and extreme numeric bounds.
/// </summary>
public class DecimalAndFinancialPrecisionTests
{
    private readonly ITestOutputHelper _output;

    public DecimalAndFinancialPrecisionTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void FinancialDecimals_PreserveExactPrecision_WithoutFloatingPointDrift()
    {
        decimal a = 0.1m;
        decimal b = 0.2m;
        decimal expectedSum = 0.3m;
        decimal largeAmount = 999_999_999_999.99m;
        decimal negativePenny = -0.01m;
        decimal tinyFraction = 0.00000001m;

        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue(a);
            sheet["A2"].SetValue(b);
            sheet["A3"].SetValue(expectedSum);
            sheet["A4"].SetValue(largeAmount);
            sheet["A5"].SetValue(negativePenny);
            sheet["A6"].SetValue(tinyFraction);

            wb.Save(ms);
        }

        ms.Position = 0;
        using (var reloaded = ExcelWorkbook.Open(ms))
        {
            var sheet = reloaded.Worksheets[0];
            Assert.Equal(a, sheet["A1"].GetDecimal());
            Assert.Equal(b, sheet["A2"].GetDecimal());
            Assert.Equal(expectedSum, sheet["A3"].GetDecimal());
            Assert.Equal(largeAmount, sheet["A4"].GetDecimal());
            Assert.Equal(negativePenny, sheet["A5"].GetDecimal());
            Assert.Equal(tinyFraction, sheet["A6"].GetDecimal());
        }

        _output.WriteLine("Financial decimal fidelity and exact precision verified successfully.");
    }

    [Fact]
    public void StreamingWorkbook_DecimalPrecision_MaintainsAccuracy()
    {
        decimal[] testDecimals = { 10.25m, 1000000.99m, -499.50m, 0.0001m };

        using var ms = new MemoryStream();
        using (var sw = ExcelStreamingWorkbook.Create(ms))
        {
            var sheet = sw.AddWorksheet("Decimals");
            for (int i = 0; i < testDecimals.Length; i++)
            {
                int row = i + 1;
                decimal d = testDecimals[i];
                sheet.WriteRow(row, rw => rw.Write(1, d));
            }
            sheet.Complete();
        }

        ms.Position = 0;
        using (var reloaded = ExcelWorkbook.Open(ms))
        {
            var sheet = reloaded.Worksheets[0];
            for (int i = 0; i < testDecimals.Length; i++)
            {
                decimal actual = sheet[i + 1, 1].GetDecimal() ?? (decimal)sheet[i + 1, 1].GetDouble()!.Value;
                Assert.Equal(testDecimals[i], actual);
            }
        }

        _output.WriteLine("Streaming decimal precision verified across multiple values.");
    }
}
