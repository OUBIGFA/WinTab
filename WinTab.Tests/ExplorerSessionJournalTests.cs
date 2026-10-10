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

    private static Task<ExplorerSession?> Request(Fixture fixture, CancellationToken token = default, bool requested = true) =>
        (Task<ExplorerSession?>)fixture.Invoke("AwaitClosedSessionsAsync", token, requested)!;

    private static Task Promotion() => WithJournal(async (fixture, saved, live) =>
    {
        var ended = Group("crash", 20, new ExplorerSessionOwner(Environment.ProcessId, 1));
        await saved.SaveAsync(Group("old-close", 10));
        await live.SaveAsync(ended);
        fixture.Invoke("PromoteEndedLiveSession");
        Check.That(saved.Snapshot!.HasSameTabs(ended), "An ended owner promotes its complete order and active tab.");
        Check.That(saved.Snapshot.EndedWithExplorer, "The promoted group records that it was open when Explorer ended.");
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
            Check.That(!copy.EndedWithExplorer, "A journal entry is not a group recovered from an ended Explorer.");
            Check.That(await live.SaveAsync(original with { Owner = null, EndedWithExplorer = true }), "A recovered group must persist.");
            using var recovered = new ExplorerSessionStore(path);
            Check.That(recovered.Snapshot!.EndedWithExplorer, "Being recovered from an ended Explorer survives a restart.");
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
