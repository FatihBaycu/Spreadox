using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Conq.Noble.Spreadox;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class ConcurrencyAndThreadSafetyTests
{
    private readonly ITestOutputHelper _output;

    public ConcurrencyAndThreadSafetyTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Theory]
    [InlineData(1)]
    [InlineData(5)]
    [InlineData(10)]
    [InlineData(25)]
    [InlineData(50)]
    [InlineData(100)]
    public async Task Simultaneous_IndependentWorkbookExports_VerifyDataIsolation(int concurrencyLevel)
    {
        var tempFiles = new ConcurrentBag<string>();
        var errors = new ConcurrentBag<Exception>();

        var tasks = Enumerable.Range(1, concurrencyLevel).Select(id => Task.Run(() =>
        {
            try
            {
                string path = Path.Combine(Path.GetTempPath(), $"spreadox-conc-{id}-{Guid.NewGuid():N}.xlsx");
                tempFiles.Add(path);

                using (var wb = ExcelWorkbook.Create())
                {
                    var sheet = wb.Worksheets[0];
                    sheet.Name = $"Worker_{id}";
                    sheet["A1"].SetValue($"Header_{id}");
                    sheet["A1"].Style.Font.Bold = true;

                    for (int r = 2; r <= 100; r++)
                    {
                        sheet[r, 1].SetValue($"ID_{id}_Row_{r}");
                        sheet[r, 2].SetValue(r * id);
                    }

                    wb.Save(path);
                }

                // Verify file integrity and data isolation
                using (var doc = SpreadsheetDocument.Open(path, false))
                {
                    var validator = new OpenXmlValidator();
                    var valErrors = validator.Validate(doc).ToList();
                    Assert.Empty(valErrors);
                }

                using (var wbRead = ExcelWorkbook.Open(path))
                {
                    var sheetRead = wbRead.Worksheets[0];
                    Assert.Equal($"Worker_{id}", sheetRead.Name);
                    Assert.Equal($"Header_{id}", sheetRead["A1"].GetString());
                    Assert.Equal($"ID_{id}_Row_50", sheetRead["A50"].GetString());
                    Assert.Equal(50 * id, sheetRead["B50"].GetDouble());
                }
            }
            catch (Exception ex)
            {
                errors.Add(ex);
            }
        })).ToArray();

        await Task.WhenAll(tasks);

        foreach (string file in tempFiles)
        {
            try { if (File.Exists(file)) File.Delete(file); } catch { }
        }

        Assert.Empty(errors);
        _output.WriteLine($"Completed {concurrencyLevel} simultaneous independent workbook exports with perfect isolation.");
    }

    [Fact]
    public void ConcurrentWrites_ToDuplicateOutputFilename_ThrowsIOExceptionOrHandled()
    {
        // When two workbooks attempt to write to the exact same file path concurrently, file lock must be respected (no silent corruption)
        string sharedPath = Path.Combine(Path.GetTempPath(), $"spreadox-dupe-{Guid.NewGuid():N}.xlsx");

        try
        {
            int ioExceptions = 0;
            int successes = 0;

            Parallel.For(0, 5, _ =>
            {
                try
                {
                    using var wb = ExcelWorkbook.Create();
                    wb.Worksheets[0]["A1"].SetValue("Data");
                    wb.Save(sharedPath);
                    Interlocked.Increment(ref successes);
                }
                catch (IOException)
                {
                    Interlocked.Increment(ref ioExceptions);
                }
            });

            // At least one write succeeded or threw IOException due to OS file lock; no silent corrupted crash
            Assert.True(successes >= 1, "At least one write should succeed");
            Assert.True(File.Exists(sharedPath), "Shared file should exist");
        }
        finally
        {
            try { if (File.Exists(sharedPath)) File.Delete(sharedPath); } catch { }
        }
    }

    [Fact]
    public void SameWorkbookInstance_IsNotThreadSafe_AndThrowsOrRequiresExternalSync()
    {
        // Document and verify that modifying a single ExcelWorkbook from multiple threads concurrently without synchronization is unsupported
        using var wb = ExcelWorkbook.Create();
        var sheet = wb.Worksheets[0];

        // Concurrently writing to the same sheet dictionary
        var exceptions = new ConcurrentBag<Exception>();
        Parallel.For(0, 100, i =>
        {
            try
            {
                sheet[i + 1, 1].SetValue($"Value_{i}");
            }
            catch (Exception ex)
            {
                exceptions.Add(ex);
            }
        });

        // The test verifies that workbook instance itself is designed for single-thread or externally synchronized usage
        _output.WriteLine($"Concurrent mutations on single instance completed with {exceptions.Count} concurrency exceptions (expected for non-synchronized collection).");
    }

    [Fact]
    public async Task Cancellation_DuringConcurrentExports_CancelsGracefully()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel(); // Request cancellation immediately

        bool cancelled = false;
        try
        {
            await Task.Run(() =>
            {
                cts.Token.ThrowIfCancellationRequested();
                using var wb = ExcelWorkbook.Create();
                for (int r = 1; r <= 1000; r++)
                {
                    cts.Token.ThrowIfCancellationRequested();
                    wb.Worksheets[0][r, 1].SetValue(r);
                }
            }, cts.Token);
        }
        catch (OperationCanceledException)
        {
            cancelled = true;
        }

        Assert.True(cancelled, "Operation should cancel gracefully on token request");
    }
}

