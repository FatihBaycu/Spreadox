using System;
using System.IO;
using System.Linq;
using Conq.Noble.Spreadox;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class ExcelCompatibilityTests
{
    private readonly ITestOutputHelper _output;

    private static readonly byte[] SamplePng =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
        0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00,
        0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
        0x00, 0x00, 0x03, 0x00, 0x01, 0x18, 0xDD, 0x8D, 0xB4, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82
    };

    public ExcelCompatibilityTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Theory]
    [InlineData("NormalSheet")]
    [InlineData("With Space")]
    [InlineData("Rapor Mayıs 2026")]
    [InlineData("Özet & İstatistik")]
    [InlineData("Turkish_ĞğÜüŞşİıÖöÇç")]
    [InlineData("1234567890123456789012345678901")] // Exactly 31 characters
    public void SpecialWorksheetNames_PassOpenXmlValidation_AndRoundTrip(string sheetName)
    {
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet.Name = sheetName;
            sheet["A1"].SetValue("TestValue");
            sheet["B1"].SetValue(1234.56);
            wb.Save(ms);
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(doc).ToList();
            Assert.Empty(errors);
        }

        ms.Position = 0;
        using (var reloaded = ExcelWorkbook.Open(ms))
        {
            Assert.Equal(sheetName, reloaded.Worksheets[0].Name);
            Assert.Equal("TestValue", reloaded.Worksheets[0]["A1"].GetString());
            Assert.Equal(1234.56, reloaded.Worksheets[0]["B1"].GetDouble());
        }
    }

    [Fact]
    public void FormulasAndCachedValues_PassOpenXmlValidation_AndRoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue(10);
            sheet["A2"].SetValue(20);
            sheet["A3"].SetValue(30);

            sheet.SetFormula("A4", "SUM(A1:A3)", cachedValue: 60);
            sheet.SetFormula("B4", "AVERAGE(A1:A3)", cachedValue: 20.0);

            wb.Save(ms);
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(doc).ToList();
            Assert.Empty(errors);
        }

        ms.Position = 0;
        using (var reloaded = ExcelWorkbook.Open(ms))
        {
            var sheet = reloaded.Worksheets[0];
            Assert.Equal("SUM(A1:A3)", sheet["A4"].Formula);
            Assert.Equal(60, sheet["A4"].GetInt());
            Assert.Equal("AVERAGE(A1:A3)", sheet["B4"].Formula);
            Assert.Equal(20.0, sheet["B4"].GetDouble());
        }
    }

    [Fact]
    public void CellStyles_Borders_Fills_FreezePanes_And_Pictures_ValidateOpenXml()
    {
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet.FreezePanes(rows: 2, columns: 1);

            var header = sheet["A1"];
            header.SetValue("Styled Header");
            header.Style.Font.Bold = true;
            header.Style.Font.Size = 14;
            header.Style.Fill.Pattern = ExcelFillPattern.Solid;
            header.Style.Fill.ForegroundColor = ExcelColor.FromRgb(40, 80, 160);
            header.Style.Font.Color = ExcelColor.White;
            header.Style.Alignment.Horizontal = ExcelHorizontalAlignment.Center;
            header.Style.Borders.Bottom.Style = ExcelBorderStyle.Medium;
            header.Style.Borders.Bottom.Color = ExcelColor.Black;

            var numCell = sheet["B2"];
            numCell.SetValue(1500000.75);
            numCell.Style.NumberFormat = NumberFormats.Currency;

            var pctCell = sheet["C2"];
            pctCell.SetValue(0.8525);
            pctCell.Style.NumberFormat = NumberFormats.PercentTwoDecimals;

            sheet.Pictures.Add(SamplePng, "A4:C8");

            wb.Save(ms);
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        Assert.Empty(errors);
    }
}
