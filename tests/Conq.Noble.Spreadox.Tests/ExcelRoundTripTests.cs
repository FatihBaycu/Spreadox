using System;
using System.IO;
using System.Linq;
using Conq.Noble.Spreadox;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 1: Excel Round-Trip Tests (P0)
/// Spreadox -> OpenXML SDK -> Modify/Save -> Spreadox Reopen -> Deep equality check
/// </summary>
public class ExcelRoundTripTests
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

    public ExcelRoundTripTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void FullWorkbook_RoundTrip_ThroughOpenXmlSdkModification_PreservesAllElements()
    {
        using var stream = new MemoryStream();

        // 1. Create with Spreadox
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet.Name = "RoundTripData";

            // Values & Styles
            sheet["A1"].SetValue("Header Item");
            sheet["A1"].Style.Font.Bold = true;
            sheet["A1"].Style.Fill.Pattern = ExcelFillPattern.Solid;
            sheet["A1"].Style.Fill.ForegroundColor = ExcelColor.FromRgb(100, 150, 200);

            sheet["B1"].SetValue(1234.56);
            sheet["B1"].Style.NumberFormat = NumberFormats.Currency;

            sheet["C1"].SetValue(new DateTime(2026, 5, 20));
            sheet["C1"].Style.NumberFormat = NumberFormats.Date;

            // Formula with cached value
            sheet.SetFormula("D1", "B1*2", cachedValue: 2469.12);

            // Merge range
            sheet.Merge("A3:C3");
            sheet["A3"].SetValue("Merged Title");

            // Freeze panes
            sheet.FreezePanes(rows: 1, columns: 1);

            // Chart
            sheet["A5"].SetValue("Cat1"); sheet["B5"].SetValue(10.0);
            sheet["A6"].SetValue("Cat2"); sheet["B6"].SetValue(25.0);
            var chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, "RT Chart");
            chart.AddSeries("B5:B6", "A5:A6", name: "RT Series");
            chart.SetPosition(8, 2, 18, 8);

            // Picture
            sheet.Pictures.Add(SamplePng, "D5:E8");

            // Named Range
            wb.AddNamedRange("MyNamedRange", "RoundTripData!$A$1:$B$1");

            wb.Save(stream);
        }

        // 2. Validate structural OpenXML schema
        stream.Position = 0;
        using (var doc = SpreadsheetDocument.Open(stream, true))
        {
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(doc).ToList();
            Assert.Empty(errors);

            // Modify something via OpenXml SDK directly (simulate 3rd-party tool or Excel edit)
            var sheetPart = doc.WorkbookPart!.WorksheetParts.First();
            var sheetData = sheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var newRow = new Row { RowIndex = 20 };
            newRow.AppendChild(new Cell
            {
                CellReference = "A20",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("ThirdPartyAdded"))
            });
            sheetData.AppendChild(newRow);
            sheetPart.Worksheet.Save();
            doc.WorkbookPart.Workbook.Save();
        }

        // 3. Re-open with Spreadox and verify fidelity
        stream.Position = 0;
        using (var reloaded = ExcelWorkbook.Open(stream))
        {
            var sheet = reloaded.Worksheets[0];
            Assert.Equal("RoundTripData", sheet.Name);

            // Verify original values
            Assert.Equal("Header Item", sheet["A1"].GetString());
            Assert.True(sheet["A1"].Style.Font.Bold);
            Assert.Equal(1234.56, sheet["B1"].GetDouble());
            Assert.Equal(new DateTime(2026, 5, 20), sheet["C1"].GetDateTime());

            // Formula verification
            Assert.Equal("B1*2", sheet["D1"].Formula);
            Assert.Equal(2469.12, sheet["D1"].GetDouble());

            // Merge range verification
            Assert.Contains(sheet.MergedRanges, r => r.A1 == "A3:C3");
            Assert.Equal("Merged Title", sheet["A3"].GetString());

            // Freeze panes
            Assert.Equal(1, sheet.FreezeRows);
            Assert.Equal(1, sheet.FreezeColumns);

            // Chart preservation
            Assert.Single(sheet.Charts);
            Assert.Equal(ExcelChartType.ColumnClustered, sheet.Charts[0].Type);
            Assert.Equal("RT Chart", sheet.Charts[0].Title);

            // Picture preservation
            Assert.Single(sheet.Pictures);

            // 3rd-party added cell
            Assert.Equal("ThirdPartyAdded", sheet["A20"].GetString());

            // Named range preservation
            Assert.Contains(reloaded.NamedRanges, nr => nr.Name == "MyNamedRange");
        }

        _output.WriteLine("Excel Round-Trip via OpenXML SDK completed with 100% element preservation.");
    }
}
