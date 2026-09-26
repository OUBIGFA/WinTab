using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;
using WinTab.Hooks;
using WinTab.Models;

internal static class ExplorerSessionLocationTests
{
    private const string ThisPc = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
    private const string RecycleBin = "shell:::{645FF040-5081-101B-9F08-00AA002F954E}";
    private const string Gallery = "shell:::{E88865EA-0E1C-4E20-9AA6-EDCD0212C87C}";

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("session locations recognize only Explorer start pages as a normal launch", StartupPagesAreExplicit);
        yield return ("session locations restore built-in local pages without probing them", BuiltInPagesNeedNoProbe);
        yield return ("session locations never probe network web relative or arbitrary shell targets", UnsafeLocationsAreNotProbed);
        yield return ("session locations keep duplicate available indexes and skip unavailable folders", AvailableIndexesRemainDistinct);
        yield return ("session locations perform real local directory checks", LocalDirectoriesAreChecked);
        yield return ("session locations follow a junction that stays on a local drive", LocalJunctionIsFollowed);
        yield return ("session locations never follow links to shares, volumes or process-relative paths", RemoteLinkTargetsAreRejected);
        yield return ("session locations keep paths confirmed before a stalled probe", ConfirmedPathsSurviveStall);
        yield return ("session locations bound a stalled device probe without queuing more work", SlowProbeIsBounded);
        yield return ("session locations honour cancellation before starting filesystem work", CancelledProbeDoesNotStart);
        yield return ("session locations can cancel a pending probe without blocking the caller", PendingProbeCanBeCancelled);
    }

    private static Task StartupPagesAreExplicit()
    {
        foreach (var location in new[]
        {
            ThisPc, "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}",
            "shell:::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}", "shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}"
        })
            Check.That(ExplorerSessionLocationPolicy.IsStartPage(location), "Known Home/This PC forms must permit a normal launch.");
        foreach (var location in new[] { "", @"C:\Users\Documents", @"\\server\share", "shell:Downloads", "shell:::{unknown}", RecycleBin, Gallery })
            Check.That(!ExplorerSessionLocationPolicy.IsStartPage(location), "An explicit or unknown folder must not be guessed to be a normal launch.");
        return Task.CompletedTask;
    }

    private static async Task BuiltInPagesNeedNoProbe()
    {
        var reads = new List<string>();
        var policy = new ExplorerSessionLocationPolicy(location => { reads.Add(location); return false; });
        var available = await policy.FindAvailableAsync(new ExplorerSession
        {
            Locations = ["::{645FF040-5081-101B-9F08-00AA002F954E}", Gallery, ThisPc, "shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"]
        }, CancellationToken.None);
        Check.That(available.SetEquals([0, 1, 2]), "Recycle Bin, Gallery and This PC are local built-in pages; Network is not.");
        Check.Equal(0, reads.Count, "Built-in pages must not reach the filesystem probe.");
    }

    private static async Task UnsafeLocationsAreNotProbed()
    {
        var reads = new List<string>();
        var policy = new ExplorerSessionLocationPolicy(location => { reads.Add(location); return true; });
        var session = new ExplorerSession
        {
            Locations = [@"\\server\share", "https://example.invalid/folder", "shell:AppsFolder", @"relative\folder",
                @"\\?\C:\folder", @"C:relative", @"C:\folder:stream", "file://server/share", @"C:\*.txt", ThisPc]
        };
        var available = await policy.FindAvailableAsync(session, CancellationToken.None);
        Check.That(available.SetEquals([9]), "Only the allowlisted shell page may bypass a local filesystem check.");
        Check.Equal(0, reads.Count, "Network or extension-backed inputs must never reach a filesystem probe.");
    }

    private static async Task AvailableIndexesRemainDistinct()
    {
        var policy = new ExplorerSessionLocationPolicy(location => location.Equals(@"C:\same", StringComparison.OrdinalIgnoreCase));
        var available = await policy.FindAvailableAsync(new ExplorerSession
        {
            Locations = [@"C:\same", @"D:\missing", "file:///C:/same", ThisPc]
        }, CancellationToken.None);
        Check.That(available.SetEquals([0, 2, 3]), "Availability must preserve duplicate saved positions and normalized file URLs.");
    }

    private static async Task LocalDirectoriesAreChecked()
    {
        var directory = NewDirectory();
        try
        {
            Directory.CreateDirectory(directory);
            var file = Path.Combine(directory, "a-file.txt");
            File.WriteAllText(file, "test");
            var policy = new ExplorerSessionLocationPolicy();
            var available = await policy.FindAvailableAsync(new ExplorerSession
            {
                Locations = [directory, file, Path.Combine(directory, "missing")]
            }, CancellationToken.None);
            Check.That(available.SetEquals([0]), "Only an existing ordinary directory is automatically restorable; files and missing paths are not.");
        }
        finally { Recycle(directory); }
    }

    private static async Task LocalJunctionIsFollowed()
    {
        var directory = NewDirectory();
        try
        {
            Directory.CreateDirectory(directory);
            var target = Path.Combine(directory, "target");
            var link = Path.Combine(directory, "junction");
            Directory.CreateDirectory(Path.Combine(target, "sub"));
            // A directory junction requires no symbolic-link privilege on Windows.
            using var process = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("cmd.exe")
            {
                Arguments = $"/d /c mklink /J \"{link}\" \"{target}\"", UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true
            })!;
            await process.WaitForExitAsync();
            Check.Equal(0, process.ExitCode, "The test junction must be created before checking the policy.");
            var policy = new ExplorerSessionLocationPolicy();
            var available = await policy.FindAvailableAsync(new ExplorerSession
            {
                Locations = [target, link, Path.Combine(link, "sub"), Path.Combine(link, "missing")]
            }, CancellationToken.None);
            Check.That(available.SetEquals([0, 1, 2]), "A junction to a local drive is an ordinary local folder; a missing child is not.");
        }
        finally { Recycle(directory); }
    }

    private static Task RemoteLinkTargetsAreRejected()
    {
        const string link = @"C:\Users\someone\link";
        foreach (var target in new[]
        {
            @"\\server\share\folder", @"\\?\UNC\server\share", @"\??\UNC\server\share", @"\\?\Volume{12345678-1234-1234-1234-123456789012}\",
            @"\??\Volume{12345678-1234-1234-1234-123456789012}\", @"\folder", @"D:relative", @"\\.\PhysicalDrive0"
        })
            Check.That(ExplorerSessionLocationPolicy.ResolveLinkTarget(link, target) == null, $"Link target {target} must not be followed.");
        Check.Equal(@"D:\data", ExplorerSessionLocationPolicy.ResolveLinkTarget(link, @"\??\D:\data"), "An NT drive path is local.");
        Check.Equal(@"D:\data", ExplorerSessionLocationPolicy.ResolveLinkTarget(link, @"D:\data"), "A drive path is local.");
        Check.Equal(@"C:\Users\other", ExplorerSessionLocationPolicy.ResolveLinkTarget(link, @"..\other"),
            "A relative link resolves against the folder that contains it.");
        return Task.CompletedTask;
    }

    private static async Task ConfirmedPathsSurviveStall()
    {
        using var gate = new ManualResetEventSlim();
        var exited = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var reads = new List<string>();
        var policy = new ExplorerSessionLocationPolicy(location =>
        {
            lock (reads) reads.Add(location);
            if (location != @"C:\slow")
                return true;
            gate.Wait();
            exited.TrySetResult();
            return true;
        });
        try
        {
            var available = await policy.FindAvailableAsync(new ExplorerSession
            {
                Locations = [@"C:\fast", ThisPc, @"C:\slow", @"C:\later"]
            }, CancellationToken.None, 300);
            Check.That(available.SetEquals([0, 1]), "Paths confirmed before the stall and built-in pages stay restorable; unconfirmed ones are skipped.");
            lock (reads)
                Check.That(!reads.Contains(@"C:\later"), "Nothing may be probed behind a stalled path within the budget.");
        }
        finally { gate.Set(); await exited.Task.WaitAsync(TimeSpan.FromSeconds(2)); }
    }

    private static async Task SlowProbeIsBounded()
    {
        using var gate = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var exited = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var reads = 0;
        var policy = new ExplorerSessionLocationPolicy(_ =>
        {
            Interlocked.Increment(ref reads);
            entered.TrySetResult();
            gate.Wait();
            exited.TrySetResult();
            return true;
        });
        var session = new ExplorerSession { Locations = [@"C:\slow", ThisPc] };
        try
        {
            var first = policy.FindAvailableAsync(session, CancellationToken.None, 40);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(2));
            Check.That((await first).SetEquals([1]), "Timed-out unchecked paths must not be passed to shell PIDL parsing.");
            for (var i = 0; i < 10; i++)
                Check.That((await policy.FindAvailableAsync(session, CancellationToken.None, 40)).SetEquals([1]),
                    "Subsequent launches may use only known shell pages while the probe is still blocked.");
            Check.Equal(1, reads, "A timeout must not spawn an unbounded number of blocked filesystem calls.");
        }
        finally { gate.Set(); await exited.Task.WaitAsync(TimeSpan.FromSeconds(2)); }
    }

    private static async Task CancelledProbeDoesNotStart()
    {
        var reads = 0;
        var policy = new ExplorerSessionLocationPolicy(_ => { Interlocked.Increment(ref reads); return true; });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await ExpectCancellation(() => policy.FindAvailableAsync(new ExplorerSession { Locations = [@"C:\not-requested"] }, cancellation.Token));
        Check.Equal(0, reads, "Cancellation before a request must prevent its filesystem operation entirely.");
    }

    private static async Task PendingProbeCanBeCancelled()
    {
        using var gate = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var exited = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var policy = new ExplorerSessionLocationPolicy(_ =>
        {
            entered.TrySetResult(); gate.Wait(); exited.TrySetResult(); return true;
        });
        using var cancellation = new CancellationTokenSource();
        try
        {
            var result = policy.FindAvailableAsync(new ExplorerSession { Locations = [@"C:\slow"] }, cancellation.Token);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(2));
            cancellation.Cancel();
            await ExpectCancellation(() => result.WaitAsync(TimeSpan.FromSeconds(1)));
        }
        finally { gate.Set(); await exited.Task.WaitAsync(TimeSpan.FromSeconds(2)); }
    }

    private static async Task ExpectCancellation(Func<Task> action)
    {
        try { await action(); }
        catch (OperationCanceledException) { return; }
        throw new InvalidOperationException("The location request must expose cancellation.");
    }

    private static string NewDirectory() => Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"));
    private static void Recycle(string directory)
    {
        if (Directory.Exists(directory)) FileSystem.DeleteDirectory(directory, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
    }
}
