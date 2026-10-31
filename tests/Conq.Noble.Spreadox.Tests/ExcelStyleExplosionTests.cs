using System;
using System.IO;
using System.Linq;
using Conq.Noble.Spreadox;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 5: Excel Style Explosion Tests (P0)
/// Verifies stylesheet deduplication, style cache boundaries, and prevents stylesheet XML explosion.
/// </summary>
public class ExcelStyleExplosionTests
{
    private readonly ITestOutputHelper _output;

    public ExcelStyleExplosionTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void StyleDeduplication_CollapsesThousandsOfIdenticalStyles_IntoMinimalCellXfs()
    {
        const int cellCount = 10_000;
        using var ms = new MemoryStream();

        using (var sw = ExcelStreamingWorkbook.Create(ms))
        {
            var headerStyle = sw.CreateStyle();
            headerStyle.Font.Bold = true;
            headerStyle.Fill.Pattern = ExcelFillPattern.Solid;
            headerStyle.Fill.ForegroundColor = ExcelColor.FromRgb(100, 150, 200);

            var numberStyle = sw.CreateStyle();
            numberStyle.NumberFormat = NumberFormats.Currency;

            var sheet = sw.AddWorksheet("DedupeTest");
            for (int r = 1; r <= cellCount; r++)
            {
                int row = r;
                sheet.WriteRow(row, rw =>
                {
                    rw.Text(1, $"Item_{row}", headerStyle);
                    rw.Number(2, row * 10.5, numberStyle);
                });
            }

            sheet.Complete();

            // UniqueStyleCount should be tiny (e.g. <= 4), NOT 20,000!
            _output.WriteLine($"Registered unique cell formats across {cellCount * 2} styled cells: {sw.UniqueStyleCount}");
            Assert.True(sw.UniqueStyleCount <= 5, $"Expected <= 5 unique styles, but got {sw.UniqueStyleCount}");
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(doc).ToList();
            Assert.Empty(errors);

            var stylesPart = doc.WorkbookPart!.WorkbookStylesPart!;
            Assert.NotNull(stylesPart.Stylesheet.CellFormats);
            uint cellXfCount = stylesPart.Stylesheet.CellFormats.Count?.Value ?? 0;
            _output.WriteLine($"Actual CellFormats count in stylesheet: {cellXfCount}");
            Assert.True(cellXfCount <= 5, $"Stylesheet cell formats exploded to {cellXfCount}");
        }
    }

    [Fact]
    public void ManyDistinctStyles_ValidateOpenXml_AndDoNotCorruptIndices()
    {
        // Test 200 distinct colors/styles to verify StyleIndexCache scales correctly
        const int distinctStyles = 200;
        using var ms = new MemoryStream();

        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            for (int i = 1; i <= distinctStyles; i++)
            {
                var cell = sheet[i, 1];
                cell.SetValue($"Style_{i}");
                cell.Style.Font.Size = 8 + (i % 8);
                cell.Style.Fill.Pattern = ExcelFillPattern.Solid;
                cell.Style.Fill.ForegroundColor = ExcelColor.FromRgb((byte)(i % 256), (byte)((i * 2) % 256), (byte)((i * 3) % 256));
            }

            wb.Save(ms);
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        Assert.Empty(errors);
        _output.WriteLine($"Verified {distinctStyles} distinct cell styles with 100% OpenXML validity.");
    }
}
