using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Hooks;
using WinTab.Managers;
using WinTab.Models;
using Fixture = ExplorerTabLifetimeTests.Fixture;

internal static class ExplorerSessionJournalTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("session journal captures last-used window with automatic restore disabled", CaptureIsAlwaysOn);
        yield return ("session journal crash teardown order cannot replace last-used group", CrashOrder);
        yield return ("session journal normal close saves closed group but remaining live group stays newer", NormalClose);
        yield return ("session journal promotes ended owner and preserves newer normal close", Promotion);
        yield return ("session journal never promotes a living process owner", LivingOwner);
        yield return ("session journal metadata round-trips and rejects invalid timestamps", Metadata);
    }

    private static Task WithJournal(Func<Fixture, ExplorerSessionStore, ExplorerSessionStore, Task> test) =>
        ExplorerSessionNativeTests.WithCapture(async (fixture, saved) =>
        {
            var directory = Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"));
            using var live = new ExplorerSessionStore(Path.Combine(directory, "live.json"));
            Set(fixture, "_liveSessionStore", live);
            Set(fixture, "_journalGate", new object());
            try { await test(fixture, saved, live); }
            finally
            {
                Set(fixture, "_captureSessions", false);
                await live.FlushAsync();
                TestCleanup.DeleteDirectory(directory);
            }
        });

    private static ExplorerSession Group(string name, long at = 0, ExplorerSessionOwner? owner = null) => new()
    {
        Locations = [@"C:\" + name, @"C:\" + name + "-other"], ActiveTabIndex = 1,
        OrderVerified = true, SavedAt = at, Owner = owner
    };

    private static Task CaptureIsAlwaysOn() => WithJournal(async (fixture, saved, live) =>
    {
        Set(fixture, "_restoreTabs", false);
        using var first = new RemoteExplorerFrame(true, explorerClass: true);
        using var second = new RemoteExplorerFrame(true, explorerClass: true);
        fixture.AddBrowser(out var a, first.Tab, handle: first.Handle);
        fixture.AddBrowser(out var b, second.Tab, handle: second.Handle);
        a.Location = @"C:\first";
        b.Location = @"C:\last-used";
        fixture.SetExplorerWindows(first.Handle, second.Handle);
        fixture.Invoke("CaptureExplorerSessions");
        Get<ExplorerSessionTracker>(fixture, "_sessionTracker").MarkUsed(b.Identity, Environment.TickCount64 + 10);
        fixture.Invoke("JournalLiveSession", false);
        Check.Equal(@"C:\last-used", live.Snapshot!.Locations[0], "Last foreground use, not enumeration order, chooses the journal.");
        Check.That(live.Snapshot.Owner!.IsRunning(), "The live journal records a real process incarnation.");
        Check.That(saved.Snapshot == null, "An open window does not replace normal-close history.");
        Check.That(await live.FlushAsync(), "A live snapshot reaches durable storage without any close event.");
    });

    private static Task CrashOrder() => WithJournal(async (fixture, saved, live) =>
    {
        using var used = new RemoteExplorerFrame(true, explorerClass: true);
        using var background = new RemoteExplorerFrame(true, explorerClass: true);
        fixture.AddBrowser(out var a, used.Tab, handle: used.Handle);
        fixture.AddBrowser(out var b, background.Tab, handle: background.Handle);
        a.Location = @"C:\in-use-before-crash";
        b.Location = @"C:\background";
        fixture.SetExplorerWindows(used.Handle, background.Handle);
        fixture.Invoke("CaptureExplorerSessions");
        Get<ExplorerSessionTracker>(fixture, "_sessionTracker").MarkUsed(a.Identity, Environment.TickCount64 + 10);
        fixture.Invoke("JournalLiveSession", false);
        var journal = live.Snapshot!;
        var usedHandle = used.Handle;
        used.Dispose();
        fixture.Invoke("NotifySessionWindowDestroyed", usedHandle);
        fixture.Invoke("CompleteClosedSessions", false);
        fixture.Invoke("JournalLiveSession", false);
        Check.That(saved.Snapshot == null, "HWND teardown is not yet proof of a normal user close.");
        Check.Equal(a.Location, live.Snapshot!.Locations[0], "A dying background survivor must not replace the intact journal.");
        var backgroundHandle = background.Handle;
        background.Dispose();
        fixture.Invoke("NotifySessionWindowDestroyed", backgroundHandle);
        fixture.SetExplorerWindows();
        // Simulate the same PID having ended without terminating any real user process.
        var ended = journal with { Owner = new ExplorerSessionOwner(Environment.ProcessId, 1) };
        await live.SaveAsync(ended);
        Set(fixture, "_liveJournaled", ended);
        fixture.Invoke("CompleteClosedSessions", true);
        Check.Equal(a.Location, saved.Snapshot!.Locations[0], "The last-used group wins even when its HWND disappeared first.");
        Check.That(saved.Snapshot.Owner == null, "The promoted group is no longer a live-process record.");
    });

    private static Task NormalClose() => WithJournal(async (fixture, saved, live) =>
    {
        using var closing = new RemoteExplorerFrame(true, explorerClass: true);
        using var survivor = new RemoteExplorerFrame(true, explorerClass: true);
        fixture.AddBrowser(out var a, closing.Tab, handle: closing.Handle);
        fixture.AddBrowser(out var b, survivor.Tab, handle: survivor.Handle);
        a.Location = @"C:\closed-normally";
        b.Location = @"C:\still-open";
        fixture.SetExplorerWindows(closing.Handle, survivor.Handle);
        fixture.Invoke("CaptureExplorerSessions");
        var handle = closing.Handle;
        closing.Dispose();
        fixture.Invoke("NotifySessionWindowDestroyed", handle);
        fixture.SetExplorerWindows(survivor.Handle);
        await (Task)fixture.Invoke("AwaitClosedSessionsAsync", CancellationToken.None)!;
        Check.Equal(a.Location, saved.Snapshot!.Locations[0], "A settled normal close becomes the saved group.");
        Check.Equal(b.Location, live.Snapshot!.Locations[0], "The surviving window remains available for later crash recovery.");
        Check.That(live.Snapshot.SavedAt > saved.Snapshot.SavedAt, "A later shutdown must prefer the still-open group.");
    });

    private static Task Promotion() => WithJournal(async (fixture, saved, live) =>
    {
        var ended = Group("crash", 20, new ExplorerSessionOwner(Environment.ProcessId, 1));
        await saved.SaveAsync(Group("old-close", 10));
        await live.SaveAsync(ended);
        fixture.Invoke("PromoteEndedLiveSession");
        Check.That(saved.Snapshot!.HasSameTabs(ended), "An ended owner promotes its complete order and active tab.");
        await saved.SaveAsync(Group("newer-close", 30));
        fixture.Invoke("PromoteEndedLiveSession");
        Check.Equal(@"C:\newer-close", saved.Snapshot!.Locations[0], "A stale live file must never supersede a newer real close.");
    });

    private static Task LivingOwner() => WithJournal(async (fixture, saved, live) =>
    {
        await live.SaveAsync(Group("still-alive", 20, ExplorerSessionOwner.Of(Environment.ProcessId)));
        fixture.Invoke("PromoteEndedLiveSession");
        Check.That(saved.Snapshot == null, "Restarting WinTab beside a living Explorer is not a crash.");
    });

    private static async Task Metadata()
    {
        var directory = Path.Combine(Path.GetTempPath(), "WinTab.Tests", Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "live.json");
        try
        {
            using var live = new ExplorerSessionStore(path);
            var owner = ExplorerSessionOwner.Of(Environment.ProcessId)!;
            var original = Group("metadata", DateTime.UtcNow.Ticks, owner);
            Check.That(await live.SaveAsync(original), "Metadata must persist.");
            using var reopened = new ExplorerSessionStore(path);
            var copy = reopened.Snapshot!;
            Check.Equal(owner, copy.Owner!, "Process incarnation survives a disk round-trip.");
            Check.Equal(original.SavedAt, copy.SavedAt, "Save ordering survives a restart.");
            Check.That(copy.HasSameTabs(original), "Order, duplicate locations and selection survive serialization.");
            Check.Throws<System.Text.Json.JsonException>(() => (original with { SavedAt = long.MaxValue }).ValidatedCopy(), "Unsupported timestamps cannot overflow save ordering.");
            Check.Throws<System.Text.Json.JsonException>(() => (original with { Owner = new ExplorerSessionOwner(-1, 1) }).ValidatedCopy(), "Invalid owners cannot be mistaken for ended processes.");
        }
        finally { TestCleanup.DeleteDirectory(directory); }
    }

    private static T Get<T>(Fixture fixture, string name) =>
        (T)typeof(ExplorerWatcher).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(fixture.Watcher)!;
    private static void Set(Fixture fixture, string name, object value) =>
        typeof(ExplorerWatcher).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(fixture.Watcher, value);
}
