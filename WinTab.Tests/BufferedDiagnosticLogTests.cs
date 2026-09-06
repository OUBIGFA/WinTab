using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;

internal static class BufferedDiagnosticLogTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("diagnostic bursts do not wait for disk access or grow the queue", BoundedQueueDoesNotBlock);
        yield return ("diagnostic files retain recent entries within the size limit", FileSizeIsBounded);
        yield return ("long diagnostic messages preserve valid UTF-8", LongMessagesStayValid);
        yield return ("diagnostic write failures are observable and stop further writes", WriteFailureIsReported);
    }

    private static async Task BoundedQueueDoesNotBlock()
    {
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var stream = new MemoryStream();
        var logger = new BufferedDiagnosticLog(() =>
        {
            started.Set();
            release.Wait();
            return stream;
        }, capacity: 8);
        Task<int>? burst = null;
        var responsive = false;
        try
        {
            Check.That(logger.TryWrite("first entry"), "An available logger should accept the first entry.");
            Check.That(started.Wait(2_000), "The background file writer should start.");
            burst = Task.Run(() =>
            {
                var accepted = 0;
                for (var index = 0; index < 1_000; index++)
                    if (logger.TryWrite($"entry {index}"))
                        accepted++;
                return accepted;
            });
            responsive = await Task.WhenAny(burst, Task.Delay(300)) == burst;
        }
        finally
        {
            release.Set();
            if (burst != null)
                await burst.WaitAsync(TimeSpan.FromSeconds(2));
            await logger.CompleteAsync().WaitAsync(TimeSpan.FromSeconds(2));
        }

        Check.That(responsive, "Logging must not make the window hook wait for disk access.");
        Check.That(burst!.Result <= 8, "The pending queue must remain bounded while the disk is blocked.");
        Check.Equal(1_000L - burst.Result, logger.DroppedMessages, "Rejected entries should be counted.");
        Check.That(Encoding.UTF8.GetString(stream.ToArray()).Contains("first entry"), "Accepted entries should be written.");
        Check.That(!logger.TryWrite("after shutdown"), "A completed logger must not accept more work.");
    }

    private static async Task FileSizeIsBounded()
    {
        var stream = new MemoryStream();
        stream.Write(Encoding.UTF8.GetBytes(new string('x', 1_024)));
        var logger = new BufferedDiagnosticLog(() => stream, capacity: 64, maximumFileBytes: 256);
        for (var index = 0; index < 30; index++)
            Check.That(logger.TryWrite($"event {index}"), "A small burst should fit in the queue.");
        Check.That(logger.TryWrite("latest event"), "The last entry should be queued.");
        await logger.CompleteAsync().WaitAsync(TimeSpan.FromSeconds(2));

        var bytes = stream.ToArray();
        Check.That(bytes.Length <= 256, "Logging must also bound a file left oversized by an earlier version.");
        Check.That(Encoding.UTF8.GetString(bytes).Contains("latest event"), "The newest entry should survive rollover.");
    }

    private static async Task LongMessagesStayValid()
    {
        var stream = new MemoryStream();
        var logger = new BufferedDiagnosticLog(() => stream, maximumFileBytes: 256);
        Check.That(logger.TryWrite("中文😀" + new string('界', 5_000)), "A long message should be safely shortened, not rejected.");
        await logger.CompleteAsync().WaitAsync(TimeSpan.FromSeconds(2));

        var bytes = stream.ToArray();
        var text = new UTF8Encoding(false, true).GetString(bytes);
        Check.That(bytes.Length <= 256, "One large message must not bypass the file size limit.");
        Check.That(text.Contains("中文😀"), "Multibyte text must remain readable.");
    }

    private static async Task WriteFailureIsReported()
    {
        var logger = new BufferedDiagnosticLog(() => throw new IOException("disk unavailable"));
        logger.TryWrite("entry");
        await logger.CompleteAsync().WaitAsync(TimeSpan.FromSeconds(2));

        Check.That(logger.LastError is IOException, "A write failure must remain observable.");
        Check.That(!logger.TryWrite("later entry"), "A failed writer must not accumulate more entries.");
    }
}
