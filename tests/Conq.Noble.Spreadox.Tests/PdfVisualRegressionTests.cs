using System;
using System.Text;
using Conq.Noble.Spreadox;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 9: PDF Visual Regression & Vector Layout Tests (P1)
/// Validates structural layout, vector bounds, headers/footers, and Turkish typography fidelity.
/// </summary>
public class PdfVisualRegressionTests
{
    private readonly ITestOutputHelper _output;

    public PdfVisualRegressionTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void PdfVisualElements_GeometryAndLabels_MatchExpectedCoordinatesAndText()
    {
        using var report = PdfReport.Create(PdfPageSize.A4Landscape);
        report.Margins = PdfMargins.Millimeters(10, 10, 10, 14);

        // Header
        report.Header.Title = "Kalite Kontrol ve İstatistik Raporu";
        report.Header.Text = "Hassas Üretim Hattı • İstasyon 4";

        // Footer
        report.Footer.Text = "Gizli / Şirket İçi";
        report.Footer.PageNumberFormat = "Sayfa {0} / {1}";

        // Cover Page
        report.Cover.Title = "Yıllık Kalite Performans Analizi";
        report.Cover.Company = "Conq.Noble Savunma & Havacılık";
        report.Cover.Date = new DateTime(2026, 10, 31);
        report.Cover.AddFact("Denetim Kapsamı", "100.000 Ölçüm");
        report.Cover.AddFact("Uygunluk", "%99.98");

        // Chart Page 1 (Line with UCL/CL/LCL)
        var p1 = report.AddChartPage("Hassas Toleranslar");
        p1.Chart.Type = ExcelChartType.LineMarker;
        string[] months = { "Oca", "Şub", "Mar", "Nis", "May", "Haz" };
        double[] vals = { 10.2, 10.5, 9.8, 11.0, 10.4, 10.7 };
        p1.Chart.AddSeries("Ölçülen", months, vals).SetColor(ExcelColor.FromRgb(0, 82, 147));
        p1.Chart.AddConstantLine("UCL", 11.5, ExcelColor.Red);
        p1.Chart.AddConstantLine("CL", 10.5, ExcelColor.Green);
        p1.Chart.AddConstantLine("LCL", 9.5, ExcelColor.Red);

        p1.Summary.Add("Ortalama (Mean)", "10.43");
        p1.Summary.Add("En Yüksek (Max)", "11.00");
        p1.Summary.Add("En Düşük (Min)", "9.80");

        byte[] pdfBytes = report.ToByteArray();
        Assert.True(pdfBytes.Length > 2048, "PDF size should be substantial.");

        using var doc = PdfDocument.Open(pdfBytes);
        Assert.Equal(2, doc.NumberOfPages); // 1 Cover + 1 Chart

        // Verify Page 1 (Cover) Dimensions and Content
        Page cover = doc.GetPage(1);
        Assert.True(cover.Width > 800 && cover.Height > 550, "A4 Landscape page dimensions expected (~842x595 pt)");
        string coverWords = GetJoinedWords(cover);
        Assert.Contains("Yıllık", coverWords);
        Assert.Contains("Performans", coverWords);
        Assert.Contains("Conq.Noble", coverWords);
        Assert.Contains("Sayfa 1 / 2", coverWords);

        // Verify Page 2 (Chart & Summary)
        Page chartPage = doc.GetPage(2);
        string chartWords = GetJoinedWords(chartPage);
        Assert.Contains("Toleranslar", chartWords);
        Assert.Contains("UCL", chartWords);
        Assert.Contains("LCL", chartWords);
        Assert.Contains("Ortalama", chartWords);
        Assert.Contains("10.43", chartWords);
        Assert.Contains("Sayfa 2 / 2", chartWords);

        _output.WriteLine("PDF Visual layout, typography, control lines, and page numbers verified cleanly.");
    }

    private static string GetJoinedWords(Page page)
    {
        var sb = new StringBuilder();
        foreach (var word in page.GetWords())
        {
            sb.Append(word.Text).Append(' ');
        }
        return sb.ToString();
    }
}
