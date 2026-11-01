using System;
using System.Diagnostics;
using System.IO;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class MemoryPressureTests
{
    private readonly ITestOutputHelper _output;

    public MemoryPressureTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void ExportUnderArtificialMemoryPressure_ExecutesSuccessfullyWithoutLeak()
    {
        // Simulate aggressive memory pressure (128 MB simulated external pressure)
        const long simulatedPressureBytes = 128 * 1024 * 1024;
        GC.AddMemoryPressure(simulatedPressureBytes);

        try
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long initialWs = Process.GetCurrentProcess().WorkingSet64;
            int initialGen2 = GC.CollectionCount(2);

            // Execute mixed Excel (streaming 50,000 rows) + PDF reports under pressure
            for (int i = 0; i < 5; i++)
            {
                // 1. Excel Streaming under pressure
                using (var ms = new MemoryStream())
                {
                    using (var sw = ExcelStreamingWorkbook.Create(ms, new ExcelStreamingOptions { UseInlineStrings = true }))
                    {
                        var sheet = sw.AddWorksheet("PressureData");
                        for (int r = 1; r <= 10_000; r++)
                        {
                            sheet.WriteRow(r, row =>
                            {
                                row.Text(1, "Metric " + r);
                                row.Number(2, r * 1.5);
                                row.Number(3, r * 2.5);
                            });
                        }
                    }
                    Assert.True(ms.Length > 0);
                }

                // 2. Vector PDF under pressure
                using (var pdf = PdfReport.Create(PdfPageSize.A4Landscape))
                {
                    pdf.Cover.Title = "Pressure Stress Test " + i;
                    pdf.Cover.Company = "Conq.Noble";
                    var page = pdf.AddChartPage("Chart Page");
                    page.Chart.Type = ExcelChartType.Line;
                    page.Chart.AddSeries("Series 1", new[] { "1", "2", "3", "4", "5" }, new double[] { 10, 25, 45, 80, 120 });

                    using var pdfMs = new MemoryStream();
                    pdf.Save(pdfMs);
                    Assert.True(pdfMs.Length > 0);
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long finalWs = Process.GetCurrentProcess().WorkingSet64;
            long wsDeltaMb = (finalWs - initialWs) / (1024 * 1024);
            int gen2Delta = GC.CollectionCount(2) - initialGen2;

            _output.WriteLine($"Memory Pressure Test: WS Delta={wsDeltaMb} MB, Gen2 Collections={gen2Delta}");
            Assert.True(wsDeltaMb < 150, $"Working set grew unexpectedly under memory pressure: {wsDeltaMb} MB");
        }
        finally
        {
            GC.RemoveMemoryPressure(simulatedPressureBytes);
        }
    }

    [Fact]
    public void ConsecutiveInterleavedRuns_MemoryScalabilityRemainsBounded()
    {
        GC.Collect();
        long baselineAllocated = GC.GetAllocatedBytesForCurrentThread();

        // Run 25 consecutive cycles of paired Excel + PDF reports
        for (int cycle = 1; cycle <= 25; cycle++)
        {
            using var excelMs = new MemoryStream();
            using (var sw = ExcelStreamingWorkbook.Create(excelMs, new ExcelStreamingOptions()))
            {
                var sheet = sw.AddWorksheet("Data");
                for (int r = 1; r <= 2000; r++)
                {
                    sheet.WriteRow(r, row =>
                    {
                        row.Text(1, "Val " + r);
                        row.Number(2, r);
                    });
                }
            }

            using var pdfMs = new MemoryStream();
            using (var report = PdfReport.Create(PdfPageSize.A4Landscape))
            {
                var page = report.AddChartPage("Col");
                page.Chart.Type = ExcelChartType.ColumnClustered;
                page.Chart.AddSeries("Vals", new[] { "A", "B", "C", "D", "E" }, new double[] { 1, 2, 3, 4, 5 });
                report.Save(pdfMs);
            }
        }

        long endAllocated = GC.GetAllocatedBytesForCurrentThread();
        long totalAllocatedMb = (endAllocated - baselineAllocated) / (1024 * 1024);
        _output.WriteLine($"25 Interleaved cycles total allocated on thread: {totalAllocatedMb} MB (Avg {(double)totalAllocatedMb / 25:F2} MB/cycle)");

        Assert.True(totalAllocatedMb < 250, $"Total thread allocations exceeded expected bounds: {totalAllocatedMb} MB");
    }
}
