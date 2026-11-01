using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class ExcelDifferentialTestingTests
{
    private readonly ITestOutputHelper _output;

    public ExcelDifferentialTestingTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void Differential_DomVsSaxVsOpenXmlSdk_ValuesAndStylesMatch()
    {
        // 1. Generate via Spreadox DOM
        using var domWb = ExcelWorkbook.Create();
        var domSheet = domWb.Worksheets[0];
        domSheet.Name = "DifferentialData";
        domSheet["A1"].SetValue("Product");
        domSheet["A1"].Style.Font.Bold = true;
        domSheet["B1"].SetValue("Revenue");
        domSheet["B1"].Style.Font.Bold = true;

        domSheet["A2"].SetValue("Alpha");
        domSheet["B2"].SetValue(1500.50m);
        domSheet["B2"].Style.NumberFormat = NumberFormats.Currency;

        domSheet["A3"].SetValue("Beta");
        domSheet["B3"].SetValue(2500.75m);
        domSheet["B3"].Style.NumberFormat = NumberFormats.Currency;

        domSheet.SetFormula("B4", "SUM(B2:B3)", cachedValue: 4001.25m);
        domSheet.FreezePanes(1, 0);
        domSheet.Merge("C1:D2");

        using var domMs = new MemoryStream();
        domWb.Save(domMs);
        domMs.Position = 0;

        // 2. Generate same semantics via Spreadox Streaming (SAX)
        using var saxMs = new MemoryStream();
        using (var sw = ExcelStreamingWorkbook.Create(saxMs, new ExcelStreamingOptions { UseInlineStrings = true }))
        {
            var sSheet = sw.AddWorksheet("DifferentialData");
            sSheet.WriteRow(1, row =>
            {
                row.Text(1, "Product");
                row.Text(2, "Revenue");
            });
            sSheet.WriteRow(2, row =>
            {
                row.Text(1, "Alpha");
                row.Number(2, 1500.50);
            });
            sSheet.WriteRow(3, row =>
            {
                row.Text(1, "Beta");
                row.Number(2, 2500.75);
            });
            sSheet.WriteRow(4, row =>
            {
                row.Number(2, 4001.25);
            });
        }
        saxMs.Position = 0;

        // 3. Inspect DOM output with OpenXML SDK
        using (var docDom = SpreadsheetDocument.Open(domMs, false))
        {
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(docDom).ToList();
            Assert.Empty(errors);

            var part = docDom.WorkbookPart!.WorksheetParts.First();
            var sheetData = part.Worksheet.Elements<SheetData>().First();
            var rows = sheetData.Elements<Row>().ToList();
            Assert.Equal(4, rows.Count);

            // Verify formula in B4
            var cellB4 = rows[3].Elements<Cell>().First(c => c.CellReference == "B4");
            Assert.Equal("SUM(B2:B3)", cellB4.CellFormula?.Text);
            Assert.Equal("4001.25", cellB4.CellValue?.Text);

            // Verify merged cells in DOM
            var mergeCells = part.Worksheet.Elements<MergeCells>().FirstOrDefault();
            Assert.NotNull(mergeCells);
            Assert.Contains(mergeCells!.Elements<MergeCell>(), mc => mc.Reference?.Value == "C1:D2");
        }

        // 4. Inspect SAX output with OpenXML SDK
        using (var docSax = SpreadsheetDocument.Open(saxMs, false))
        {
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(docSax).ToList();
            Assert.Empty(errors);

            var part = docSax.WorkbookPart!.WorksheetParts.First();
            var sheetData = part.Worksheet.Elements<SheetData>().First();
            var rows = sheetData.Elements<Row>().ToList();
            Assert.Equal(4, rows.Count);

            // Verify value in B4
            var cellB4 = rows[3].Elements<Cell>().First(c => c.CellReference == "B4");
            Assert.Equal("4001.25", cellB4.CellValue?.Text);
        }

        // 5. Cross-Reload: Reload DOM generated package into fresh Spreadox ExcelWorkbook
        domMs.Position = 0;
        using var reloadedDom = ExcelWorkbook.Open(domMs);
        var rSheet = reloadedDom.Worksheets[0];
        Assert.NotNull(rSheet);
        Assert.Equal("Product", rSheet["A1"].GetString());
        Assert.True(rSheet["A1"].Style.Font.Bold);
        Assert.Equal("Alpha", rSheet["A2"].GetString());
        Assert.Equal(1500.50m, rSheet["B2"].GetDecimal());
        Assert.Equal("Beta", rSheet["A3"].GetString());
        Assert.Equal(2500.75m, rSheet["B3"].GetDecimal());
        Assert.Equal("SUM(B2:B3)", rSheet["B4"].Formula);
        Assert.Equal(4001.25m, rSheet["B4"].GetDecimal());

        _output.WriteLine("Differential verification passed across DOM, SAX, and OpenXML SDK.");
    }

    [Fact]
    public void Differential_RoundTripWithThirdPartyMutation_PreservesDrawingMLAndImages()
    {
        byte[] originalBytes;
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet.Name = "ChartData";
            sheet["A1"].SetValue("Month");
            sheet["B1"].SetValue("Sales");
            sheet["A2"].SetValue("Jan"); sheet["B2"].SetValue(100);
            sheet["A3"].SetValue("Feb"); sheet["B3"].SetValue(200);
            sheet["A4"].SetValue("Mar"); sheet["B4"].SetValue(300);

            var chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, "Monthly Sales");
            chart.AddSeries("B2:B4", "A2:A4", name: "Sales");
            chart.SetPosition(5, 1, 15, 8);

            // 1x1 transparent PNG
            byte[] pngBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");
            sheet.Pictures.Add(pngBytes, "A16:D20");

            using var ms = new MemoryStream();
            wb.Save(ms);
            originalBytes = ms.ToArray();
        }

        // Simulate 3rd party OpenXML tool (adding a new cell and saving)
        byte[] mutatedBytes;
        using (var editMs = new MemoryStream())
        {
            editMs.Write(originalBytes, 0, originalBytes.Length);
            editMs.Position = 0;
            using (var doc = SpreadsheetDocument.Open(editMs, true))
            {
                var wsPart = doc.WorkbookPart!.WorksheetParts.First();
                var sheetData = wsPart.Worksheet.Elements<SheetData>().First();
                var newRow = new Row { RowIndex = 25 };
                newRow.AppendChild(new Cell { CellReference = "A25", DataType = CellValues.String, CellValue = new CellValue("AppendedByThirdPartyTool") });
                sheetData.AppendChild(newRow);
                doc.Save();
            }
            mutatedBytes = editMs.ToArray();
        }

        // Reload mutated package in Spreadox
        using var verifyMs = new MemoryStream(mutatedBytes);
        using var reloaded = ExcelWorkbook.Open(verifyMs);
        var rSheet = reloaded.Worksheets[0];
        Assert.NotNull(rSheet);

        // Verify original data
        Assert.Equal("Jan", rSheet["A2"].GetString());
        Assert.Equal(100, rSheet["B2"].GetInt());
        Assert.Equal("Mar", rSheet["A4"].GetString());
        Assert.Equal(300, rSheet["B4"].GetInt());

        // Verify appended data
        Assert.Equal("AppendedByThirdPartyTool", rSheet["A25"].GetString());

        // Verify charts and pictures preserved
        Assert.Single(rSheet.Charts);
        Assert.Equal("Monthly Sales", rSheet.Charts[0].Title);
        Assert.Single(rSheet.Pictures);

        _output.WriteLine("Third-party round-trip differential verification successful.");
    }
}
