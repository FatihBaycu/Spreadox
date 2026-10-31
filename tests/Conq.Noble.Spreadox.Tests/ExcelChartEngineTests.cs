using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Conq.Noble.Spreadox;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class ExcelChartEngineTests
{
    private readonly ITestOutputHelper _output;

    public ExcelChartEngineTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Theory]
    [InlineData(1_000)]
    [InlineData(10_000)]
    public void LineChart_WithManyPoints_WritesAndValidates(int pointCount)
    {
        var sw = Stopwatch.StartNew();
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("Index");
            sheet["B1"].SetValue("Reading");

            for (int i = 1; i <= pointCount; i++)
            {
                sheet[i + 1, 1].SetValue(i);
                sheet[i + 1, 2].SetValue(Math.Sin(i * 0.05) * 10.0);
            }

            var chart = sheet.Charts.Add(ExcelChartType.Line, $"Line_{pointCount}");
            chart.AddSeries($"B2:B{pointCount + 1}", $"A2:A{pointCount + 1}", name: "SinWave");
            chart.SetPosition(2, 4, 20, 15);

            wb.Save(ms);
        }

        sw.Stop();
        _output.WriteLine($"Line chart ({pointCount:N0} points): {sw.ElapsedMilliseconds} ms, size={ms.Length:N0} bytes");

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        Assert.Empty(errors);

        var chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();
        Assert.NotNull(chartPart.ChartSpace);
        var lineChart = chartPart.ChartSpace.Descendants<LineChart>().FirstOrDefault();
        Assert.NotNull(lineChart);
    }

    [Fact]
    public void ScatterChart_WithManyPoints_WritesAndValidates()
    {
        const int pointCount = 5_000;
        var sw = Stopwatch.StartNew();
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("X");
            sheet["B1"].SetValue("Y");

            for (int i = 1; i <= pointCount; i++)
            {
                sheet[i + 1, 1].SetValue(i * 0.1);
                sheet[i + 1, 2].SetValue(i * i * 0.01);
            }

            var chart = sheet.Charts.Add(ExcelChartType.Scatter, "Scatter5k");
            chart.AddSeries($"B2:B{pointCount + 1}", $"A2:A{pointCount + 1}", name: "Parabola");
            chart.SetPosition(2, 4, 20, 15);

            wb.Save(ms);
        }

        sw.Stop();
        _output.WriteLine($"Scatter chart ({pointCount:N0} points): {sw.ElapsedMilliseconds} ms, size={ms.Length:N0} bytes");

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void ChartEngine_HandlesNull_Zero_Negative_And_DecimalValues()
    {
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("Category");
            sheet["B1"].SetValue("Val");

            sheet["A2"].SetValue("NullVal");
            // B2 left empty (null)
            sheet["A3"].SetValue("Zero");
            sheet["B3"].SetValue(0.0);
            sheet["A4"].SetValue("Negative");
            sheet["B4"].SetValue(-42.75);
            sheet["A5"].SetValue("Decimal");
            sheet["B5"].SetValue(0.00001234567);

            var chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, "EdgeCases");
            chart.AddSeries("B2:B5", "A2:A5", name: "Metrics");
            chart.SetPosition(2, 4, 18, 12);

            wb.Save(ms);
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void PieAndDoughnut_WithLargeCategoryCounts()
    {
        const int categories = 50;
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("Slice");
            sheet["B1"].SetValue("Value");

            for (int i = 1; i <= categories; i++)
            {
                sheet[i + 1, 1].SetValue($"Category_{i}");
                sheet[i + 1, 2].SetValue(100.0 / categories);
            }

            var pie = sheet.Charts.Add(ExcelChartType.Pie, "LargePie");
            pie.AddSeries($"B2:B{categories + 1}", $"A2:A{categories + 1}", name: "Distribution");
            pie.SetPosition(2, 4, 15, 12);

            var doughnut = sheet.Charts.Add(ExcelChartType.Doughnut, "LargeDoughnut");
            doughnut.DoughnutHoleSize = 65;
            doughnut.AddSeries($"B2:B{categories + 1}", $"A2:A{categories + 1}", name: "Distribution");
            doughnut.SetPosition(2, 14, 15, 22);

            wb.Save(ms);
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void DateBasedAxes_And_MultipleAxes_StructureValid()
    {
        using var ms = new MemoryStream();
        using (var wb = ExcelWorkbook.Create())
        {
            var sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("Date");
            sheet["B1"].SetValue("Primary");
            sheet["C1"].SetValue("Secondary %");

            var baseDate = new DateTime(2026, 1, 1);
            for (int i = 1; i <= 10; i++)
            {
                sheet[i + 1, 1].SetValue(baseDate.AddDays(i));
                sheet[i + 1, 1].Style.NumberFormat = NumberFormats.Date;
                sheet[i + 1, 2].SetValue(i * 100);
                sheet[i + 1, 3].SetValue(i * 0.05);
            }

            var combo = sheet.Charts.Add(ExcelChartType.Combo, "DateCombo");
            combo.ValueAxis.Title = "Volume";
            combo.SecondaryValueAxis.Title = "Percentage";
            combo.AddSeries("B2:B11", "A2:A11", name: "Volume").AsChartType(ExcelChartType.ColumnClustered);
            combo.AddSeries("C2:C11", "A2:A11", name: "Percentage").AsChartType(ExcelChartType.Line).OnSecondaryAxis();
            combo.SetPosition(2, 5, 20, 15);

            wb.Save(ms);
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        Assert.Empty(errors);
    }
}
