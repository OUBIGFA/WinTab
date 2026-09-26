using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        yield return ("a diagnostic line that repeats the previous one is counted instead of written again", RepeatedLinesAreCollapsed);
        yield return ("the explorer diagnostic log creates the folder it is configured to write to", LogFolderIsCreated);
        yield return ("the daily log keeps today's and yesterday's files and deletes older ones", DailyLogKeepsTwoDays);
        yield return ("the daily log starts a new file when the date changes while it runs", DailyLogRollsOverAtMidnight);
        yield return ("the daily log bounds a day that outgrows its size", DailyLogBoundsOneDay);
        yield return ("the daily log recovers after its folder becomes unwritable", DailyLogRecoversAfterFailure);
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
        var text = Encoding.UTF8.GetString(stream.ToArray());
        Check.That(text.Contains("first entry"), "Accepted entries should be written.");
        Check.That(text.Contains($"diagnostic queue dropped {logger.DroppedMessages} messages"),
            "The file must report missing entries after the writer resumes, not just count them in memory.");
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

    private static async Task RepeatedLinesAreCollapsed()
    {
        var stream = new MemoryStream();
        var logger = new BufferedDiagnosticLog(() => stream);
        for (var index = 0; index < 5; index++)
            logger.TryWrite("Selection capture unavailable");
        logger.TryWrite("Navigation click captured");
        logger.TryWrite("Navigation click captured");
        await logger.CompleteAsync().WaitAsync(TimeSpan.FromSeconds(2));

        var lines = Encoding.UTF8.GetString(stream.ToArray())
            .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries).Select(line => line[13..]).ToArray();
        Check.That(lines.SequenceEqual(["Selection capture unavailable", "(previous line repeated 4 more times)",
                "Navigation click captured", "(previous line repeated 1 more time)"]),
            $"A recurring line is written once with its count, and the count is kept at shutdown: {string.Join(" | ", lines)}");
    }

    private static Task LogFolderIsCreated()
    {
        // WINTAB_LOG_FILE may name a folder that does not exist yet; logging must still start instead of
        // failing silently on the first write.
        var path = Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"), "logs", "explorer-debug.log");
        using (var stream = WinTab.Hooks.ExplorerDebugLog.OpenLogFile(path))
            stream.Write(Encoding.UTF8.GetBytes("entry"));
        Check.Equal("entry", File.ReadAllText(path), "The log file must be written inside the folder it was configured with.");
        return Task.CompletedTask;
    }

    private static string NewLogFolder() => Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"), "logs");

    private static void DeleteTestFolder(string folder)
    {
        var root = Path.GetDirectoryName(folder)!;
        if (Directory.Exists(root))
            Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(root,
                Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
    }

    private static Task DailyLogKeepsTwoDays()
    {
        var folder = NewLogFolder();
        try
        {
            var today = new DateTime(2026, 9, 26, 10, 0, 0);
            Directory.CreateDirectory(folder);
            foreach (var name in new[] { "WinTab-20260925.log", "WinTab-20260924.log", "WinTab-20260924.1.log", "WinTab-20260901.log",
                         "notes.txt", "WinTab-backup.log", "Other-20260101.log" })
                File.WriteAllText(Path.Combine(folder, name), "old");
            using (var sink = new DailyLogSink(folder))
                sink.Append(today, Encoding.UTF8.GetBytes("today\n"));
            var names = Directory.GetFiles(folder).Select(Path.GetFileName).Order(StringComparer.Ordinal).ToArray();
            Check.That(names.SequenceEqual(["Other-20260101.log", "WinTab-20260925.log", "WinTab-20260926.log", "WinTab-backup.log", "notes.txt"]),
                $"Only today's and yesterday's log files remain, and files the log did not write are untouched: {string.Join(", ", names)}");
            Check.Equal("today\n", File.ReadAllText(Path.Combine(folder, "WinTab-20260926.log")), "Today's lines go to today's file.");
        }
        finally { DeleteTestFolder(folder); }
        return Task.CompletedTask;
    }

    private static Task DailyLogRollsOverAtMidnight()
    {
        var folder = NewLogFolder();
        try
        {
            using (var sink = new DailyLogSink(folder))
            {
                sink.Append(new DateTime(2026, 9, 24, 23, 59, 0), Encoding.UTF8.GetBytes("first\n"));
                sink.Append(new DateTime(2026, 9, 25, 0, 1, 0), Encoding.UTF8.GetBytes("second\n"));
                sink.Append(new DateTime(2026, 9, 26, 0, 1, 0), Encoding.UTF8.GetBytes("third\n"));
            }
            var names = Directory.GetFiles(folder).Select(Path.GetFileName).Order(StringComparer.Ordinal).ToArray();
            Check.That(names.SequenceEqual(["WinTab-20260925.log", "WinTab-20260926.log"]),
                $"A WinTab that runs for days keeps only the last two days as well: {string.Join(", ", names)}");
            Check.Equal("second\n", File.ReadAllText(Path.Combine(folder, "WinTab-20260925.log")), "Each line is in its own day's file.");
        }
        finally { DeleteTestFolder(folder); }
        return Task.CompletedTask;
    }

    private static Task DailyLogBoundsOneDay()
    {
        var folder = NewLogFolder();
        try
        {
            var day = new DateTime(2026, 9, 26, 12, 0, 0);
            var line = Encoding.UTF8.GetBytes(new string('x', 300) + "\n");
            using (var sink = new DailyLogSink(folder, maximumFileBytes: 1_024))
                for (var index = 0; index < 20; index++)
                    sink.Append(day, line);
            var current = new FileInfo(Path.Combine(folder, "WinTab-20260926.log"));
            var previous = new FileInfo(Path.Combine(folder, "WinTab-20260926.1.log"));
            Check.That(current.Exists && current.Length <= 1_024 && previous.Exists && previous.Length <= 1_024,
                "A day keeps its newest lines and one earlier file, each within the size limit.");
            Check.Equal(2, Directory.GetFiles(folder).Length, "A busy day never grows beyond two files.");
        }
        finally { DeleteTestFolder(folder); }
        return Task.CompletedTask;
    }

    private static Task DailyLogRecoversAfterFailure()
    {
        var root = NewLogFolder();
        // A file where the log folder should be makes the folder impossible to create.
        Directory.CreateDirectory(Path.GetDirectoryName(root)!);
        File.WriteAllText(root, "blocking file");
        try
        {
            long now = 0;
            using (var sink = new DailyLogSink(root, getTickCount: () => now))
            {
                var day = new DateTime(2026, 9, 26, 12, 0, 0);
                sink.Append(day, Encoding.UTF8.GetBytes("lost\n"));
                Check.That(sink.LastError != null && sink.DroppedLines == 1, "A line that cannot be written is counted and its error kept.");
                sink.Append(day, Encoding.UTF8.GetBytes("lost during backoff\n"));
                File.Move(root, root + "-blocked");
                now = 29_999;
                sink.Append(day, Encoding.UTF8.GetBytes("still in backoff\n"));
                Check.That(!Directory.Exists(root), "Storage failures must be backed off even after the folder becomes writable.");
                now = 30_000;
                sink.Append(day.AddSeconds(30), Encoding.UTF8.GetBytes("recovered\n"));
                Check.Equal(3L, sink.DroppedLines, "All lost lines, including the backoff interval, must be counted.");
            }
            var text = File.ReadAllText(Path.Combine(root, "WinTab-20260926.log"));
            Check.That(text.Contains("recovered"), "Logging must actually resume after the storage failure clears.");
            Check.That(text.Contains("dropped 3 lines") && text.Contains("IOException"),
                "The recovered log must explain the missing interval with a count and the storage failure.");
        }
        finally { DeleteTestFolder(root); }
        return Task.CompletedTask;
    }
}
