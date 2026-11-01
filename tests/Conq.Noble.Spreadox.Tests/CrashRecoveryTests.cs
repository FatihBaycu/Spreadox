using System;
using System.IO;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class CrashRecoveryTests
{
    private readonly ITestOutputHelper _output;

    public CrashRecoveryTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void AbruptInterruption_IncompleteFileCannotBeOpenedAsValid()
    {
        string incompletePath = Path.Combine(Path.GetTempPath(), $"spreadox_crash_{Guid.NewGuid():N}.xlsx");
        try
        {
            // Simulate sudden crash: write 500 rows and intentionally abort before disposing workbook
            var fs = File.Create(incompletePath);
            var sw = ExcelStreamingWorkbook.Create(fs, new ExcelStreamingOptions { UseInlineStrings = true });
            var sheet = sw.AddWorksheet("CrashData");
            for (int r = 1; r <= 500; r++)
            {
                sheet.WriteRow(r, row =>
                {
                    row.Text(1, $"Row {r}");
                    row.Number(2, r);
                });
            }

            // Simulate process abrupt termination (killing stream without disposing the packaging writer)
            // DO NOT call sw.Dispose();
            // Force stream closure without writing central directory
            fs.Flush();
            fs.Close();

            // Verification: Opening this incomplete file MUST fail and never be treated as valid
            Assert.ThrowsAny<Exception>(() =>
            {
                using var verifyWb = ExcelWorkbook.Open(incompletePath);
            });

            _output.WriteLine("Verified: Incomplete export file was safely rejected upon open.");
        }
        finally
        {
            if (File.Exists(incompletePath))
            {
                File.Delete(incompletePath);
            }
        }
    }

    [Fact]
    public void FaultInjectionDuringStreaming_UnderlyingResourcesAreCleanedUp()
    {
        string faultyPath = Path.Combine(Path.GetTempPath(), $"spreadox_fault_{Guid.NewGuid():N}.xlsx");
        try
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                // Create with path -> ownsStream = true
                using var sw = ExcelStreamingWorkbook.Create(faultyPath);
                var sheet = sw.AddWorksheet("FaultSheet");
                for (int r = 1; r <= 100; r++)
                {
                    sheet.WriteRow(r, row =>
                    {
                        row.Text(1, "Test " + r);
                        row.Number(2, r * 10);
                    });

                    if (r == 10)
                    {
                        // Simulate logic / process crash: non-monotonic row index
                        sheet.WriteRow(5, row => row.Text(1, "Out of order"));
                    }
                }
            });

            // Confirm that file is not locked and can be safely deleted or opened
            Assert.True(File.Exists(faultyPath));
            File.Delete(faultyPath);
            _output.WriteLine("Verified: Exception during streaming properly cleaned up file locks and resources.");
        }
        finally
        {
            if (File.Exists(faultyPath))
            {
                File.Delete(faultyPath);
            }
        }
    }

    [Fact]
    public void DestinationFileLock_PreventsSilentOverwritesOrCorruptions()
    {
        string lockedPath = Path.Combine(Path.GetTempPath(), $"spreadox_locked_{Guid.NewGuid():N}.xlsx");
        try
        {
            // Create an initial valid file
            using (var initialWb = ExcelWorkbook.Create())
            {
                initialWb.Worksheets[0].SetValue("A1", "Pre-existing content");
                initialWb.Save(lockedPath);
            }

            // Lock the file exclusively with another process/stream
            using (var lockStream = new FileStream(lockedPath, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                // Attempting to export to locked file should fail cleanly with IOException rather than silently corrupting
                Assert.Throws<IOException>(() =>
                {
                    using var secondWb = ExcelWorkbook.Create();
                    secondWb.Worksheets[0].SetValue("A1", "New content");
                    secondWb.Save(lockedPath);
                });
            }

            // Verify original file remained intact
            using (var verifyWb = ExcelWorkbook.Open(lockedPath))
            {
                Assert.Equal("Pre-existing content", verifyWb.Worksheets[0]["A1"].GetString());
            }

            _output.WriteLine("Verified: Existing destination file was protected from corruption during lock collision.");
        }
        finally
        {
            if (File.Exists(lockedPath))
            {
                File.Delete(lockedPath);
            }
        }
    }

    private sealed class ResourceTrackingStream : MemoryStream
    {
        public bool InjectFailure { get; set; }
        public bool IsDisposed { get; private set; }

        public override void Write(byte[] buffer, int offset, int count)
        {
            if (InjectFailure)
            {
                throw new IOException("Simulated fatal I/O failure during streaming write.");
            }
            base.Write(buffer, offset, count);
        }

        protected override void Dispose(bool disposing)
        {
            IsDisposed = true;
            base.Dispose(disposing);
        }
    }
}
