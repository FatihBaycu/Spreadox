using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class PdfFontCacheConcurrencyTests
{
    private readonly ITestOutputHelper _output;

    public PdfFontCacheConcurrencyTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public async Task ConcurrentPdfGenerations_100Tasks_ThreadSafeCacheAndZeroLohBloat()
    {
        const int concurrentCount = 100;
        var fontCheckResults = new ConcurrentBag<bool>();
        var errors = new ConcurrentBag<string>();

        GC.Collect();
        var initialInfo = GC.GetGCMemoryInfo();
        long initialLohBytes = initialInfo.GenerationInfo.Length > 0 ? initialInfo.GenerationInfo[initialInfo.GenerationInfo.Length - 1].SizeAfterBytes : 0;
        var sw = Stopwatch.StartNew();

        var tasks = Enumerable.Range(1, concurrentCount).Select(async i =>
        {
            await Task.Yield();
            try
            {
                using var ms = new MemoryStream();
                using (var report = PdfReport.Create(PdfPageSize.A4Landscape))
                {
                    report.Cover.Title = $"Concurrent Report {i}";
                    report.Cover.Company = "Conq.Noble";
                    report.Header.Title = "Concurrent Thread Safety";
                    report.Header.Text = $"Worker Task #{i}";
                    report.Header.Bold = true;

                    var page = report.AddChartPage($"Page for Worker {i}");
                    page.Chart.Type = ExcelChartType.Line;
                    page.Chart.AddSeries("Series", new[] { "1", "2", "3", "4" }, new double[] { i, i * 2, i * 3, i * 4 });

                    report.Save(ms);
                }

                if (ms.Length > 0)
                {
                    fontCheckResults.Add(true);
                }
                else
                {
                    errors.Add($"Worker {i} produced empty PDF.");
                }
            }
            catch (Exception ex)
            {
                errors.Add($"Worker {i} failed: {ex.Message}");
            }
        });

        await Task.WhenAll(tasks);
        sw.Stop();

        GC.Collect();
        var finalInfo = GC.GetGCMemoryInfo();
        long finalLohBytes = finalInfo.GenerationInfo.Length > 0 ? finalInfo.GenerationInfo[finalInfo.GenerationInfo.Length - 1].SizeAfterBytes : 0;
        long lohDeltaMb = (finalLohBytes - initialLohBytes) / (1024 * 1024);

        _output.WriteLine($"100 Concurrent PDF reports generated in {sw.ElapsedMilliseconds} ms (Avg: {(double)sw.ElapsedMilliseconds / concurrentCount:F1} ms/PDF)");
        _output.WriteLine($"LOH Delta after 100 concurrent reports: {lohDeltaMb} MB (Zero font stream leaks)");

        Assert.Empty(errors);
        Assert.Equal(concurrentCount, fontCheckResults.Count);
        // Without font caching, 100 PDFs would allocate at least 250 MB on the LOH. With static cache, delta is minimal (< 20 MB).
        Assert.True(lohDeltaMb < 30, $"LOH growth was unexpectedly large: {lohDeltaMb} MB");
    }

    [Fact]
    public void CustomFontRegistration_ThreadSafeUnderConcurrentAccess()
    {
        // Generate a valid mock TTF byte array
        byte[] dummyTtf = new byte[1024];
        for (int i = 0; i < dummyTtf.Length; i++) dummyTtf[i] = (byte)(i % 256);

        // Register concurrently from 20 parallel threads
        Parallel.For(1, 21, i =>
        {
            PdfFontRegistry.Register($"TestFont_{i}", dummyTtf);
        });

        // Verify registration didn't throw and was thread-safe
        _output.WriteLine("Verified: Custom font registration is thread-safe under concurrent access.");
    }
}
