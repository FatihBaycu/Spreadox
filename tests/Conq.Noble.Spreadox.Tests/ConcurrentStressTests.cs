using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class ConcurrentStressTests
{
    private readonly ITestOutputHelper _output;

    public ConcurrentStressTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public async Task SustainedConcurrentGeneration_100Workers_MaintainsDataIsolationAndHighThroughput()
    {
        const int workerCount = 100;
        var latenciesMs = new ConcurrentBag<long>();
        var errors = new ConcurrentBag<string>();
        int initialHandles = Process.GetCurrentProcess().HandleCount;

        var sw = Stopwatch.StartNew();

        var tasks = Enumerable.Range(1, workerCount).Select(async workerId =>
        {
            await Task.Yield(); // ensure thread-pool execution
            var itemSw = Stopwatch.StartNew();
            try
            {
                // Each worker creates an Excel workbook and a PDF report
                string expectedExcelToken = $"WORKER_{workerId}_{Guid.NewGuid():N}";
                string expectedPdfToken = $"PDF_METRIC_{workerId}";

                // 1. Excel Generation
                using var excelMs = new MemoryStream();
                using (var wb = ExcelWorkbook.Create())
                {
                    var sheet = wb.Worksheets[0];
                    sheet.Name = $"WorkerSheet_{workerId}";
                    sheet["A1"].SetValue(expectedExcelToken);
                    sheet["B1"].SetValue(workerId * 100);
                    sheet.SetFormula("C1", "B1*2", cachedValue: workerId * 200);
                    wb.Save(excelMs);
                }

                // Verify Excel Isolation immediately
                excelMs.Position = 0;
                using (var verifyWb = ExcelWorkbook.Open(excelMs))
                {
                    string actualToken = verifyWb.Worksheets[0]["A1"].GetString();
                    if (actualToken != expectedExcelToken)
                    {
                        errors.Add($"Worker {workerId} data isolation mismatch! Expected {expectedExcelToken}, got {actualToken}");
                    }
                    if (verifyWb.Worksheets[0]["B1"].GetInt() != workerId * 100)
                    {
                        errors.Add($"Worker {workerId} numeric mismatch!");
                    }
                }

                // 2. PDF Generation
                using var pdfMs = new MemoryStream();
                using (var pdf = PdfReport.Create(PdfPageSize.A4Landscape))
                {
                    pdf.Cover.Title = expectedPdfToken;
                    var page = pdf.AddChartPage($"Worker {workerId} Page");
                    page.Chart.Type = ExcelChartType.ColumnClustered;
                    page.Chart.AddSeries("Data", new[] { "A", "B", "C" }, new double[] { workerId, workerId * 2, workerId * 3 });
                    pdf.Save(pdfMs);
                }

                if (pdfMs.Length == 0)
                {
                    errors.Add($"Worker {workerId} produced empty PDF!");
                }

                itemSw.Stop();
                latenciesMs.Add(itemSw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                errors.Add($"Worker {workerId} threw {ex.GetType().Name}: {ex.Message}");
            }
        });

        await Task.WhenAll(tasks);
        sw.Stop();

        int finalHandles = Process.GetCurrentProcess().HandleCount;
        int handleDelta = finalHandles - initialHandles;

        var sortedLatencies = latenciesMs.OrderBy(l => l).ToList();
        long p50 = sortedLatencies[(int)(sortedLatencies.Count * 0.50)];
        long p95 = sortedLatencies[(int)(sortedLatencies.Count * 0.95)];
        long p99 = sortedLatencies[(int)(sortedLatencies.Count * 0.99)];
        double avg = sortedLatencies.Average();

        _output.WriteLine($"100 Concurrent Workers: Completed in {sw.ElapsedMilliseconds} ms ({workerCount * 2} documents)");
        _output.WriteLine($"Latency Distribution: Avg={avg:F1}ms, p50={p50}ms, p95={p95}ms, p99={p99}ms");
        _output.WriteLine($"OS Handle Delta: {handleDelta}");

        Assert.Empty(errors);
        Assert.True(sw.ElapsedMilliseconds < 15000, $"Concurrent generation was too slow: {sw.ElapsedMilliseconds} ms");
    }
}
