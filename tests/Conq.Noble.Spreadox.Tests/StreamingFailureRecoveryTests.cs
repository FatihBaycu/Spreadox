using System;
using System.IO;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 6: Streaming Failure Recovery Tests (P0)
/// Verifies behavior when I/O fails, streams close prematurely, or errors occur mid-stream.
/// </summary>
public class StreamingFailureRecoveryTests
{
    private readonly ITestOutputHelper _output;

    public StreamingFailureRecoveryTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>Simulated stream that fails with IOException after a given byte budget (e.g. disk full).</summary>
    private sealed class FaultyStream : MemoryStream
    {
        private readonly long _maxBytesBeforeFailure;
        private long _written;

        public FaultyStream(long maxBytesBeforeFailure)
        {
            _maxBytesBeforeFailure = maxBytesBeforeFailure;
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            if (_written + count > _maxBytesBeforeFailure)
            {
                throw new IOException("Simulated disk full or I/O failure.");
            }

            base.Write(buffer, offset, count);
            _written += count;
        }
    }

    [Fact]
    public void DiskFull_MidStream_ThrowsIOException_AndDisposesGracefully()
    {
        var faulty = new FaultyStream(maxBytesBeforeFailure: 2048);
        bool ioExceptionCaught = false;

        try
        {
            using (var sw = ExcelStreamingWorkbook.Create(faulty))
            {
                var sheet = sw.AddWorksheet("Heavy");
                for (int r = 1; r <= 10_000; r++)
                {
                    int row = r;
                    sheet.WriteRow(row, rw =>
                    {
                        rw.Text(1, "LargeDataStringRow_" + row);
                        rw.Number(2, row * 999.99);
                    });
                }
                sheet.Complete();
            }
        }
        catch (IOException ex) when (ex.Message.Contains("Simulated disk full"))
        {
            ioExceptionCaught = true;
        }

        Assert.True(ioExceptionCaught, "Expected IOException on simulated disk full.");
        _output.WriteLine("Verified disk-full failure mid-stream throws cleanly without process instability.");
    }

    [Fact]
    public void NonAscendingRowIndex_ThrowsInvalidOperationException()
    {
        using var ms = new MemoryStream();
        using var sw = ExcelStreamingWorkbook.Create(ms);
        var sheet = sw.AddWorksheet("Ordered");

        sheet.WriteRow(10, rw => rw.Text(1, "Row 10"));

        // Attempting to write row 5 after row 10 must throw
        var ex = Assert.Throws<InvalidOperationException>(() =>
        {
            sheet.WriteRow(5, rw => rw.Text(1, "Row 5"));
        });

        Assert.Contains("ascending order", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DisposedWorkbook_ThrowsObjectDisposedException_OnSubsequentOperations()
    {
        var ms = new MemoryStream();
        var sw = ExcelStreamingWorkbook.Create(ms);
        sw.Dispose();

        Assert.Throws<ObjectDisposedException>(() => sw.AddWorksheet("Another"));
        Assert.Throws<ObjectDisposedException>(() => sw.CreateStyle());
    }

    [Fact]
    public void UnwritableStream_ThrowsArgumentException()
    {
        using var readOnlyStream = new MemoryStream(new byte[100], writable: false);
        Assert.Throws<ArgumentException>(() => ExcelStreamingWorkbook.Create(readOnlyStream));
    }
}
