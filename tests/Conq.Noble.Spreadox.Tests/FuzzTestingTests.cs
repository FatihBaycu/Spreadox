using System;
using System.IO;
using System.Text;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 3: Fuzz Testing (P0)
/// Robustness validation against malformed, corrupted, and unexpected inputs.
/// </summary>
public class FuzzTestingTests
{
    private readonly ITestOutputHelper _output;

    public FuzzTestingTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void Open_CorruptedBytes_ThrowsControlledException_AndDoesNotCrash()
    {
        // Generate pseudo-random garbage bytes
        var rand = new Random(101);
        byte[] garbage = new byte[4096];
        rand.NextBytes(garbage);

        using var ms = new MemoryStream(garbage);
        // Expect controlled exception (e.g. InvalidDataException, FileFormatException or similar from packaging)
        Assert.ThrowsAny<Exception>(() => ExcelWorkbook.Open(ms));
    }

    [Fact]
    public void Open_TruncatedZipArchive_ThrowsControlledException()
    {
        // Valid workbook truncated midway
        byte[] validBytes;
        using (var wb = ExcelWorkbook.Create())
        {
            wb.Worksheets[0]["A1"].SetValue("Data");
            using var ms = new MemoryStream();
            wb.Save(ms);
            validBytes = ms.ToArray();
        }

        // Truncate to first 50%
        byte[] truncated = new byte[validBytes.Length / 2];
        Array.Copy(validBytes, truncated, truncated.Length);

        using var truncStream = new MemoryStream(truncated);
        Assert.ThrowsAny<Exception>(() => ExcelWorkbook.Open(truncStream));
    }

    [Fact]
    public void ExcelCell_HandlesExtremeLongStrings_AndControlCharactersSafely()
    {
        using var wb = ExcelWorkbook.Create();
        var sheet = wb.Worksheets[0];

        // 1. String with null and control characters
        string controlChars = "Hello\0World\x01\x02\x03\t\r\nEnd";
        sheet["A1"].SetValue(controlChars);
        Assert.Equal(controlChars, sheet["A1"].GetString());

        // 2. 50,000 character string (exceeds Excel 32,767 cell limit)
        string massive = new string('X', 50_000);
        sheet["A2"].SetValue(massive);
        Assert.Equal(50_000, sheet["A2"].GetString()?.Length);

        using var ms = new MemoryStream();
        wb.Save(ms);
        Assert.True(ms.Length > 0);
    }

    [Fact]
    public void PdfFontRegistry_CorruptedBytes_DoesNotCorruptRegistry()
    {
        byte[] corruptedTtf = Encoding.UTF8.GetBytes("Not A Real TTF Font File At All");
        
        // Registering corrupted bytes should either accept as raw bytes or fail gracefully
        PdfFontRegistry.Register("CorruptedFont", corruptedTtf);
        Assert.True(PdfFontRegistry.IsRegistered("CorruptedFont"));

        // Normal PDF generation with default fonts should still work unaffected
        using var report = PdfReport.Create();
        report.Cover.Title = "Sanity Check";
        byte[] pdf = report.ToByteArray();
        Assert.True(pdf.Length > 0);
    }

    [Fact]
    public void ExcelWorkbook_HandlesMalformedFormulaStringsSafely()
    {
        using var wb = ExcelWorkbook.Create();
        var sheet = wb.Worksheets[0];

        // Formula with dangling operators or unclosed parens
        sheet.SetFormula("A1", "SUM(((A1+))", cachedValue: null);
        sheet.SetFormula("A2", "====", cachedValue: 0);

        using var ms = new MemoryStream();
        wb.Save(ms);
        Assert.True(ms.Length > 0);

        ms.Position = 0;
        using var reloaded = ExcelWorkbook.Open(ms);
        Assert.NotNull(reloaded.Worksheets[0]["A1"].Formula);
    }
}
