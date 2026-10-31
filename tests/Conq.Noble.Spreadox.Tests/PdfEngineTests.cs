using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Conq.Noble.Spreadox;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class PdfEngineTests
{
    private readonly ITestOutputHelper _output;

    public PdfEngineTests(ITestOutputHelper output)
    {
        _output = output;
    }

    private static string ExtractPageText(Page page)
    {
        var sb = new StringBuilder();
        foreach (var word in page.GetWords())
        {
            sb.Append(word.Text).Append(' ');
        }
        return sb.ToString();
    }

    [Fact]
    public void TurkishCharacters_And_Unicode_RenderAndAreSelectable()
    {
        const string turkishSample = "İçişleri Bakanlığı Özet Çizelgesi: Ğaziantep & Şanlıurfa";

        using var report = PdfReport.Create();
        report.Cover.Title = turkishSample;
        report.Cover.Company = "Conq.Noble Türkiye Ar-Ge";
        report.Header.Title = "Rapor Başlığı: Ölçüm Değerleri (Ğ, Ü, Ş, İ, Ö, Ç)";
        report.Footer.Text = "Gizli • Dağıtılamaz • Şirket İçi";

        var page = report.AddChartPage("İstasyon Değerleri");
        page.Chart.Type = ExcelChartType.LineMarker;
        page.Chart.AddSeries("Seri 1 (Açık/Kapalı)", new[] { "Ocak", "Şubat", "Mart", "Nisan", "Mayıs" }, new[] { 10.5, 20.2, 15.8, 30.1, 25.4 });
        page.Summary.Add("Ölçüm Ortalaması", "20.4");
        page.Summary.Add("En Düşük", "10.5");

        byte[] pdfBytes = report.ToByteArray();
        Assert.True(pdfBytes.Length > 0);

        using var pdfDoc = PdfDocument.Open(pdfBytes);
        Assert.Equal(2, pdfDoc.NumberOfPages);

        string coverText = ExtractPageText(pdfDoc.GetPage(1));
        _output.WriteLine("Cover page extracted: " + coverText);
        Assert.Contains("Bakanlığı", coverText);
        Assert.Contains("Çizelgesi", coverText);

        string chartText = ExtractPageText(pdfDoc.GetPage(2));
        _output.WriteLine("Chart page extracted: " + chartText);
        Assert.Contains("İstasyon", chartText);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(10)]
    [InlineData(100)]
    public void PdfReport_Pagination_And_MultiPageScaling(int chartPages)
    {
        var sw = Stopwatch.StartNew();
        using var report = PdfReport.Create();
        report.Cover.Title = $"Volume Test {chartPages} Pages";
        report.Footer.PageNumberFormat = "Sayfa {0} / {1}";

        string[] cats = { "A", "B", "C", "D", "E" };
        double[] vals = { 10, 25, 18, 32, 29 };

        for (int p = 1; p <= chartPages; p++)
        {
            var page = report.AddChartPage($"Page_{p}");
            page.Chart.Type = (p % 2 == 0) ? ExcelChartType.ColumnClustered : ExcelChartType.Line;
            page.Chart.AddSeries($"Series_{p}", cats, vals);
            page.Summary.Add("Metric", p.ToString());
        }

        byte[] pdf = report.ToByteArray();
        sw.Stop();

        _output.WriteLine($"Generated {chartPages} chart pages (+1 cover) in {sw.ElapsedMilliseconds} ms, size = {pdf.Length / 1024} KB");

        using var doc = PdfDocument.Open(pdf);
        Assert.Equal(chartPages + 1, doc.NumberOfPages);

        // Verify page numbering in footer
        string lastPageText = ExtractPageText(doc.GetPage(doc.NumberOfPages));
        Assert.Contains($"Sayfa {chartPages + 1} / {chartPages + 1}", lastPageText);
    }

    [Fact]
    public void PdfChart_HandlesNegativeAndExtremeValues()
    {
        using var report = PdfReport.Create();
        var page = report.AddChartPage("Extreme Values");
        page.Chart.Type = ExcelChartType.Line;

        string[] cats = { "Min", "Zero", "ExtremeNegative", "ExtremePositive", "SmallDecimal" };
        double[] values = { -100.0, 0.0, -999999.99, 1000000.5, 0.000001 };

        page.Chart.AddSeries("Extremes", cats, values);
        page.Chart.AddConstantLine("Upper Limit", 500000, ExcelColor.Red);
        page.Chart.AddConstantLine("Lower Limit", -500000, ExcelColor.Blue);

        byte[] pdf = report.ToByteArray();
        Assert.True(pdf.Length > 1000);

        using var doc = PdfDocument.Open(pdf);
        Assert.Equal(2, doc.NumberOfPages);
    }
}

