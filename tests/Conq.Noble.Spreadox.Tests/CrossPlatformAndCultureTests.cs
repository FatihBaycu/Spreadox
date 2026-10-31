using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Conq.Noble.Spreadox;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 11: Cross-Platform & Culture Invariance Tests (P1)
/// Ensures behavior is 100% identical regardless of host OS culture or font availability.
/// </summary>
public class CrossPlatformAndCultureTests
{
    private readonly ITestOutputHelper _output;

    public CrossPlatformAndCultureTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Theory]
    [InlineData("tr-TR")] // Turkish (dotted/dotless i, comma decimal)
    [InlineData("de-DE")] // German (comma decimal, period thousands)
    [InlineData("fr-FR")] // French (comma decimal, space thousands)
    [InlineData("en-US")] // US English
    [InlineData("ar-SA")] // Arabic
    public void NumericSerialization_IsCultureInvariant_UnderAllCultures(string cultureName)
    {
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
        CultureInfo originalUiCulture = Thread.CurrentThread.CurrentUICulture;

        try
        {
            var targetCulture = new CultureInfo(cultureName);
            Thread.CurrentThread.CurrentCulture = targetCulture;
            Thread.CurrentThread.CurrentUICulture = targetCulture;

            using var ms = new MemoryStream();
            using (var wb = ExcelWorkbook.Create())
            {
                var sheet = wb.Worksheets[0];
                sheet["A1"].SetValue(1234.5678);
                sheet["A2"].SetValue(0.0001);
                sheet["A3"].SetValue(-9876.54);

                wb.Save(ms);
            }

            // Verify OpenXml XML uses '.' decimal separator strictly
            ms.Position = 0;
            using (var doc = SpreadsheetDocument.Open(ms, false))
            {
                var validator = new OpenXmlValidator();
                var errors = validator.Validate(doc);
                Assert.Empty(errors);

                var sheetPart = doc.WorkbookPart!.WorksheetParts.First();
                var cells = sheetPart.Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);

                // XML numbers MUST have '.' and NEVER ','
                Assert.Contains(".", cells["A1"].CellValue!.Text);
                Assert.DoesNotContain(",", cells["A1"].CellValue!.Text);

                Assert.Contains(".", cells["A2"].CellValue!.Text);
                Assert.DoesNotContain(",", cells["A2"].CellValue!.Text);
            }

            // Reload and verify double value is identical
            ms.Position = 0;
            using (var reloaded = ExcelWorkbook.Open(ms))
            {
                Assert.Equal(1234.5678, reloaded.Worksheets[0]["A1"].GetDouble());
                Assert.Equal(0.0001, reloaded.Worksheets[0]["A2"].GetDouble());
                Assert.Equal(-9876.54, reloaded.Worksheets[0]["A3"].GetDouble());
            }

            _output.WriteLine($"Culture '{cultureName}' verified: numbers correctly serialized with invariant '.' decimal point.");
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = originalCulture;
            Thread.CurrentThread.CurrentUICulture = originalUiCulture;
        }
    }

    [Fact]
    public void EmbeddedFonts_PermitHeadlessPdfGeneration_WithoutSystemFonts()
    {
        // On Linux Docker containers without font packages, embedded Liberation Sans enables PDF generation
        using var report = PdfReport.Create();
        report.Cover.Title = "Docker Container Headless Verification";
        var page = report.AddChartPage("Chart 1");
        page.Chart.Type = ExcelChartType.Line;
        page.Chart.AddSeries("Metrics", new[] { "P1", "P2" }, new[] { 10.0, 20.0 });

        byte[] pdf = report.ToByteArray();
        Assert.True(pdf.Length > 1000);
    }
}
