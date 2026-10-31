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
/// Group 4: Excel Formula Compatibility & Injection Security Tests (P0)
/// </summary>
public class ExcelFormulaCompatibilityTests
{
    private readonly ITestOutputHelper _output;

    public ExcelFormulaCompatibilityTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void StandardFormulas_AreWrittenCorrectly_WithCachedValues()
    {
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            for (int i = 1; i <= 5; i++)
            {
                sheet[i, 1].SetValue(i * 10); // A1..A5: 10, 20, 30, 40, 50 (Sum = 150, Avg = 30)
                sheet[i, 2].SetValue(i);      // B1..B5: 1, 2, 3, 4, 5
            }

            sheet.SetFormula("A6", "=SUM(A1:A5)", cachedValue: 150);
            sheet.SetFormula("A7", "AVERAGE(A1:A5)", cachedValue: 30.0);
            sheet.SetFormula("A8", "=COUNT(A1:A5)", cachedValue: 5);
            sheet.SetFormula("A9", "IF(A6>100, 1, 0)", cachedValue: 1);
            sheet.SetFormula("A10", "IF(SUM(A1:A5)>100, AVERAGE(B1:B5), 0)", cachedValue: 3.0);

            wb.Save(ms);
        }

        // Validate OpenXML structure
        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(doc).ToList();
            Assert.Empty(errors);

            var wsPart = doc.WorkbookPart!.WorksheetParts.First();
            var cells = wsPart.Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);

            // Verify formulas are stored without leading '='
            Assert.Equal("SUM(A1:A5)", cells["A6"].CellFormula!.Text);
            Assert.Equal("150", cells["A6"].CellValue!.Text);

            Assert.Equal("AVERAGE(A1:A5)", cells["A7"].CellFormula!.Text);
            Assert.Equal("30", cells["A7"].CellValue!.Text);

            Assert.Equal("COUNT(A1:A5)", cells["A8"].CellFormula!.Text);
            Assert.Equal("5", cells["A8"].CellValue!.Text);

            Assert.Equal("IF(A6>100, 1, 0)", cells["A9"].CellFormula!.Text);
            Assert.Equal("IF(SUM(A1:A5)>100, AVERAGE(B1:B5), 0)", cells["A10"].CellFormula!.Text);
        }

        // Round-trip back to Spreadox
        ms.Position = 0;
        using (var reloaded = ExcelWorkbook.Open(ms))
        {
            var sheet = reloaded.Worksheets[0];
            Assert.Equal("SUM(A1:A5)", sheet["A6"].Formula);
            Assert.Equal(150, sheet["A6"].GetInt());
            Assert.Equal(30.0, sheet["A7"].GetDouble());
            Assert.Equal(5, sheet["A8"].GetInt());
            Assert.Equal(1, sheet["A9"].GetInt());
            Assert.Equal(3.0, sheet["A10"].GetDouble());
        }

        _output.WriteLine("Verified standard formulas (SUM, AVERAGE, COUNT, IF, nested) and cached value serialization.");
    }

    [Fact]
    public void FormulaInjectionSecurity_TextStartingWithEqualSign_IsNotTreatedAsFormula()
    {
        // When untrusted user input starts with '=', '+', '-', or '@', calling SetValue(string)
        // must treat it strictly as a string literal and NOT emit a <f> formula tag.
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("=cmd|'/C calc'!A0"); // Classic DDE injection payload
            sheet["A2"].SetValue("+12345");
            sheet["A3"].SetValue("-SUM(A1:A2)");
            sheet["A4"].SetValue("@SUM(1,2)");

            wb.Save(ms);
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var wsPart = doc.WorkbookPart!.WorksheetParts.First();
            var cells = wsPart.Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);

            // Ensure NONE of these cells contain a <f> CellFormula element
            Assert.Null(cells["A1"].CellFormula);
            Assert.Null(cells["A2"].CellFormula);
            Assert.Null(cells["A3"].CellFormula);
            Assert.Null(cells["A4"].CellFormula);
        }

        ms.Position = 0;
        using (var reloaded = ExcelWorkbook.Open(ms))
        {
            var sheet = reloaded.Worksheets[0];
            Assert.False(sheet["A1"].HasFormula);
            Assert.Equal("=cmd|'/C calc'!A0", sheet["A1"].GetString());
        }

        _output.WriteLine("Formula injection security confirmed: untrusted strings are never emitted as formula expressions.");
    }
}
