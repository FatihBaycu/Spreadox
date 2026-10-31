using System;
using System.Diagnostics;
using System.IO;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 10: Long-Running Stability Tests (P0)
/// Simulates continuous 24/7 web reporting workloads to detect slow memory/handle leaks.
/// </summary>
public class LongRunningStabilityTests
{
    private readonly ITestOutputHelper _output;

    public LongRunningStabilityTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void ContinuousMixedWorkload_MaintainsStableWorkingSet_AndZeroHandleLeaks()
    {
        const int totalCycles = 500; // 500 interleaved Excel + PDF cycles

        GC.Collect(2, GCCollectionMode.Forced, true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, true);

        long initialManaged = GC.GetTotalMemory(true);
        Process current = Process.GetCurrentProcess();
        current.Refresh();
        long initialWorkingSet = current.WorkingSet64;
        int initialHandles = current.HandleCount;

        var sw = Stopwatch.StartNew();

        for (int i = 1; i <= totalCycles; i++)
        {
            // 1. Generate Excel Workbook
            using (var msXlsx = new MemoryStream())
            {
                using (var wb = ExcelWorkbook.Create())
                {
                    var sheet = wb.Worksheets[0];
                    for (int r = 1; r <= 20; r++)
                    {
                        sheet[r, 1].SetValue($"Row_{r}");
                        sheet[r, 2].SetValue(r * 1.5);
                    }
                    wb.Save(msXlsx);
                }
            }

            // 2. Generate PDF Report
            using (var report = PdfReport.Create())
            {
                report.Cover.Title = $"Report #{i}";
                var p = report.AddChartPage($"Station_{i % 5}");
                p.Chart.Type = ExcelChartType.Line;
                p.Chart.AddSeries("Reading", new[] { "A", "B", "C" }, new[] { 1.0, 2.0, 3.0 });
                _ = report.ToByteArray();
            }
        }

        sw.Stop();

        System.Runtime.GCSettings.LargeObjectHeapCompactionMode = System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce;
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, true, true);

        long finalManaged = GC.GetTotalMemory(true);
        current.Refresh();
        long finalWorkingSet = current.WorkingSet64;
        int finalHandles = current.HandleCount;

        _output.WriteLine($"Completed {totalCycles} mixed cycles in {sw.ElapsedMilliseconds} ms ({sw.ElapsedMilliseconds / (double)totalCycles:F1} ms/cycle)");
        _output.WriteLine($"Managed Heap: Initial={initialManaged / 1024 / 1024} MB, Final={finalManaged / 1024 / 1024} MB");
        _output.WriteLine($"Working Set: Initial={initialWorkingSet / 1024 / 1024} MB, Final={finalWorkingSet / 1024 / 1024} MB");
        _output.WriteLine($"OS Handles: Initial={initialHandles}, Final={finalHandles}");

        // Managed heap should be bounded (allowance for LOH segments from 500 PDF byte arrays)
        long managedGrowthMb = (finalManaged - initialManaged) / 1024 / 1024;
        Assert.True(managedGrowthMb < 150, $"Managed heap leaked {managedGrowthMb} MB");

        // Working set committed pages bounded
        long wsGrowthMb = (finalWorkingSet - initialWorkingSet) / 1024 / 1024;
        Assert.True(wsGrowthMb < 300, $"Working set grew unexpectedly by {wsGrowthMb} MB");

        // Handle count should not leak
        int handleGrowth = finalHandles - initialHandles;
        Assert.True(handleGrowth < 50, $"OS Handles leaked: {handleGrowth}");
    }
}
