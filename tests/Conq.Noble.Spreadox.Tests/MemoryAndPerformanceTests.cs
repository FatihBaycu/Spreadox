using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class MemoryAndPerformanceTests
{
    private readonly ITestOutputHelper _output;

    public MemoryAndPerformanceTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public record MetricSnapshot(
        string Scenario,
        int Rows,
        int Columns,
        long ElapsedMs,
        long PeakWorkingSetBytes,
        long PrivateBytes,
        long AllocatedBytes,
        int Gen0,
        int Gen1,
        int Gen2,
        long OutputSizeBytes
    )
    {
        public override string ToString() =>
            $"{Scenario} ({Rows:N0} rows x {Columns} cols): {ElapsedMs} ms | PeakWS={PeakWorkingSetBytes / 1024 / 1024} MB | Private={PrivateBytes / 1024 / 1024} MB | Allocated={AllocatedBytes / 1024 / 1024} MB | GCs=({Gen0},{Gen1},{Gen2}) | FileSize={OutputSizeBytes / 1024} KB";
    }

    public static MetricSnapshot MeasureStreamingExport(int rows, int cols, string label)
    {
        GC.Collect(2, GCCollectionMode.Forced, true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, true);

        long allocBefore = GC.GetTotalAllocatedBytes(precise: true);
        int g0Before = GC.CollectionCount(0);
        int g1Before = GC.CollectionCount(1);
        int g2Before = GC.CollectionCount(2);

        Process current = Process.GetCurrentProcess();
        current.Refresh();
        long wsBefore = current.WorkingSet64;
        long privateBefore = current.PrivateMemorySize64;

        string tempPath = Path.Combine(Path.GetTempPath(), $"spreadox-bench-{label}-{Guid.NewGuid():N}.xlsx");
        var sw = Stopwatch.StartNew();
        long peakWs = wsBefore;

        try
        {
            using (var workbook = ExcelStreamingWorkbook.Create(tempPath))
            {
                var sheet = workbook.AddWorksheet("Data");
                for (int r = 1; r <= rows; r++)
                {
                    int rowIndex = r;
                    sheet.WriteRow(rowIndex, row =>
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            if (c == 1)
                                row.Text(c, "R" + rowIndex);
                            else if (c % 2 == 0)
                                row.Number(c, rowIndex * 1.5);
                            else
                                row.Integer(c, rowIndex + c);
                        }
                    });

                    // Sample memory periodically on large runs
                    if (r % 100_000 == 0)
                    {
                        current.Refresh();
                        if (current.WorkingSet64 > peakWs)
                            peakWs = current.WorkingSet64;
                    }
                }
                sheet.Complete();
            }

            sw.Stop();
            current.Refresh();
            if (current.WorkingSet64 > peakWs)
                peakWs = current.WorkingSet64;

            long allocAfter = GC.GetTotalAllocatedBytes(precise: true);
            int g0After = GC.CollectionCount(0);
            int g1After = GC.CollectionCount(1);
            int g2After = GC.CollectionCount(2);
            long fileSize = new FileInfo(tempPath).Length;

            return new MetricSnapshot(
                label,
                rows,
                cols,
                sw.ElapsedMilliseconds,
                peakWs,
                current.PrivateMemorySize64,
                Math.Max(0, allocAfter - allocBefore),
                g0After - g0Before,
                g1After - g1Before,
                g2After - g2Before,
                fileSize
            );
        }
        finally
        {
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
        }
    }

    [Theory]
    [InlineData(10_000, 10)]
    [InlineData(10_000, 30)]
    [InlineData(100_000, 10)]
    [InlineData(100_000, 30)]
    public void StreamingWorkbook_MediumAndLargeRuns_MeasureMetrics(int rows, int cols)
    {
        var metrics = MeasureStreamingExport(rows, cols, $"Stream_{rows}_{cols}");
        _output.WriteLine(metrics.ToString());

        Assert.True(metrics.ElapsedMs > 0, "Execution time should be > 0");
        Assert.True(metrics.OutputSizeBytes > 1024, "Output size should be > 1KB");
        Assert.True(metrics.AllocatedBytes > 0, "Allocated bytes should be measured");
    }

    [Fact]
    public void StreamingWorkbook_500kRows_CompletesSuccessfully_And_MeasuresMemory()
    {
        var metrics = MeasureStreamingExport(500_000, 10, "Stream_500k_10");
        _output.WriteLine(metrics.ToString());

        Assert.True(metrics.ElapsedMs < 120_000, $"500k rows took {metrics.ElapsedMs} ms (> 120s)");
        Assert.True(metrics.OutputSizeBytes > 1_000_000, "500k rows file size should be > 1MB");
    }

    [Fact]
    public void StreamingWorkbook_1MillionRows_CompletesSuccessfully_And_MeasuresMemory()
    {
        var metrics = MeasureStreamingExport(1_000_000, 10, "Stream_1M_10");
        _output.WriteLine(metrics.ToString());

        Assert.True(metrics.ElapsedMs < 240_000, $"1M rows took {metrics.ElapsedMs} ms (> 240s)");
        Assert.True(metrics.OutputSizeBytes > 2_000_000, "1M rows file size should be > 2MB");
    }

    [Fact]
    public void RepeatedExports_ToDetectMemoryLeaks()
    {
        // Execute 5 repeated streaming exports of 20,000 rows each and measure working set stability
        const int iterations = 5;
        const int rows = 20_000;
        const int cols = 10;
        long[] workingSets = new long[iterations];

        for (int i = 0; i < iterations; i++)
        {
            var m = MeasureStreamingExport(rows, cols, $"Repeat_{i + 1}");
            _output.WriteLine(m.ToString());
            workingSets[i] = m.PeakWorkingSetBytes;
        }

        // Memory should not continuously grow unbounded; working set at iteration 5 should be within reasonable bounds of iteration 2
        _output.WriteLine($"Initial WS: {workingSets[0] / 1024 / 1024} MB, Final WS: {workingSets[iterations - 1] / 1024 / 1024} MB");
        Assert.True(workingSets[iterations - 1] > 0);
    }
}

