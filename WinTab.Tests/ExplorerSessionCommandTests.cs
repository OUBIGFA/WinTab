using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Models;

internal static class ExplorerSessionCommandTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("session command history restores newest first without merging equal paths", HistoryOrder);
        yield return ("session command history peeks newest first and takes exact entries", PeekAndTake);
        yield return ("session command history deduplicates OnQuit and window teardown after consumption", DuplicateCallbacks);
        yield return ("session command failed reopen keeps chronological order and bounded capacity", FailedReopenOrder);
        yield return ("session command disabling history invalidates an in-flight failed reopen", ClearInvalidatesReturn);
        yield return ("session command moved live tabs are not restored", MovedTabIsSkipped);
        yield return ("session command new-window plan preserves duplicate order and selected saved index", RequestedPlan);
        yield return ("session command new-window plan skips unavailable paths and falls back to first", RequestedPlanSkips);
        yield return ("session command unavailable closed tab remains available for retry", UnavailableClosedTab);
        yield return ("session command reopen passes over newer closes that cannot be opened now", ReopenPassesOverUnavailable);
        yield return ("session command reopen stops at a folder whose drive does not answer in time", ReopenStopsAtUndecidedFolder);
        yield return ("session command serializes requests and releases the gate after failures", SerializesRequests);
        yield return ("session command independent window protection lasts through request scope", IndependentScope);
    }

    private static ClosedTab Tab(int token, string location = @"C:\same") =>
        new(location, new WindowIdentity(0, 1, 1, token), default);

    /// <summary>Takes the newest close, as a reopen of a location that is available does.</summary>
    private static bool Pop(ExplorerClosedTabHistory history, out ClosedTab tab)
    {
        tab = history.Peek().FirstOrDefault();
        return tab.Sequence != 0 && history.TryTake(tab);
    }

    private static nint[] Tokens(ExplorerClosedTabHistory history) => history.Peek().Select(tab => tab.Tab.Token).ToArray();

    private static Task PeekAndTake()
    {
        var history = new ExplorerClosedTabHistory();
        history.Push(Tab(1, @"C:\first"));
        history.Push(Tab(2, @"C:\second"));
        history.Push(Tab(3, @"C:\third"));
        var tabs = history.Peek();
        Check.That(Tokens(history).SequenceEqual([3, 2, 1]), "Closes are offered newest first.");
        Check.That(history.TryTake(tabs[1]), "An older close can be taken without the newer ones.");
        Check.That(!history.TryTake(tabs[1]), "A taken close cannot be taken twice.");
        Check.That(Tokens(history).SequenceEqual([3, 1]), "The other closes keep their order.");
        history.Return(tabs[1]);
        Check.That(Tokens(history).SequenceEqual([3, 2, 1]), "A failed reopen returns the close to its place.");
        history.Clear();
        history.Push(Tab(4));
        Check.That(!history.TryTake(tabs[0]), "Disabling history forgets closes that were already peeked.");
        Check.Equal(1, history.Count, "A stale take must not remove a close recorded afterwards.");
        return Task.CompletedTask;
    }

    private static Task HistoryOrder()
    {
        var history = new ExplorerClosedTabHistory(2);
        history.Push(Tab(1, @"C:\evicted"));
        history.Push(Tab(2));
        history.Push(Tab(3));
        Check.Equal(2, history.Count, "The history is bounded.");
        Check.That(Pop(history, out var newest) && newest.Tab.Token == 3, "Latest closure wins.");
        Check.That(Pop(history, out var older) && older.Tab.Token == 2, "Equal paths from distinct tabs survive separately.");
        Check.That(!Pop(history, out _), "Evicted entries do not return.");
        Check.Throws<ArgumentOutOfRangeException>(() => new ExplorerClosedTabHistory(0), "An empty capacity is invalid.");
        return Task.CompletedTask;
    }

    private static Task DuplicateCallbacks()
    {
        var history = new ExplorerClosedTabHistory();
        history.Push(Tab(1));
        history.Push(Tab(1));
        Check.Equal(1, history.Count, "OnQuit and frame teardown describe only one close.");
        Check.That(Pop(history, out _), "The user consumes the close.");
        history.Push(Tab(1));
        Check.Equal(0, history.Count, "A delayed teardown must not reinsert an already reopened tab.");
        history.Push(Tab(2));
        Check.Equal(1, history.Count, "Recycled HWNDs with a new token remain distinct.");
        return Task.CompletedTask;
    }

    private static Task FailedReopenOrder()
    {
        var history = new ExplorerClosedTabHistory(2);
        history.Push(Tab(1));
        Check.That(Pop(history, out var pending), "A restore starts.");
        history.Push(Tab(2));
        history.Return(pending);
        history.Return(pending);
        Check.Equal(2, history.Count, "A retry is not duplicated.");
        Check.That(Pop(history, out var newer) && newer.Tab.Token == 2, "A later user close stays newest.");
        Check.That(Pop(history, out var returned) && returned.Tab.Token == 1, "Failed entries retain their place.");
        history.Push(Tab(3));
        history.Push(Tab(4));
        history.Return(returned);
        Check.Equal(2, history.Count, "Returning an old failure never exceeds the cap.");
        return Task.CompletedTask;
    }

    private static Task ClearInvalidatesReturn()
    {
        var history = new ExplorerClosedTabHistory();
        history.Push(Tab(1));
        Check.That(Pop(history, out var pending), "A restore starts.");
        history.Clear();
        history.Return(pending);
        Check.Equal(0, history.Count, "Disabling history forgets pending restores too.");
        return Task.CompletedTask;
    }

    private static Task MovedTabIsSkipped() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        var history = new ExplorerClosedTabHistory();
        history.Push(Tab(1));
        history.Push(new ClosedTab(@"C:\moved", WindowIdentity.Capture(fixture.Window.FirstTab), default));
        Check.That(Pop(history, out var tab) && tab.Tab.Token == 1, "A tab still alive after tear-off was not closed.");
        Check.That(!Pop(history, out _), "The moved tab is discarded without opening a duplicate.");
        return Task.CompletedTask;
    });

    private static Task RequestedPlan()
    {
        var session = new ExplorerSession { Locations = [@"C:\same", @"C:\other", @"C:\same"], ActiveTabIndex = 2, OrderVerified = true };
        var (plan, first) = ExplorerSessionRestorePlan.CreateForNewWindow(session, new HashSet<int> { 0, 1, 2 })!.Value;
        Check.Equal(0, first, "The new window starts at the first saved location, not a placeholder.");
        Check.That(plan.Tabs.Select(tab => tab.SavedIndex).SequenceEqual([1, 2]), "All other tabs retain saved order and duplicates.");
        Check.Equal(2, plan.ActiveSavedIndex, "The active duplicate is selected by saved index.");
        Check.That(!plan.CloseInitialTab && !plan.ActivateInitialTab, "The first saved tab is never closed.");
        return Task.CompletedTask;
    }

    private static Task RequestedPlanSkips()
    {
        var session = new ExplorerSession { Locations = [@"C:\missing", @"C:\available", @"C:\other"], ActiveTabIndex = 0, OrderVerified = true };
        var (plan, first) = ExplorerSessionRestorePlan.CreateForNewWindow(session, new HashSet<int> { 1 })!.Value;
        Check.Equal(1, first, "The first available path supplies the new window.");
        Check.Equal(2, plan.SkippedCount, "Unavailable entries are reported.");
        Check.That(plan.ActivateInitialTab && plan.Tabs.Length == 0, "A single survivor needs no extra tab.");
        Check.That(ExplorerSessionRestorePlan.CreateForNewWindow(session, new HashSet<int>()) == null, "No available path opens no window.");
        return Task.CompletedTask;
    }

    private static Task UnavailableClosedTab() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        Set(fixture.Watcher, "_captureSessions", true);
        Set(fixture.Watcher, "_recordClosedTabs", true);
        Set(fixture.Watcher, "_sessionLocationPolicy", new ExplorerSessionLocationPolicy(_ => false));
        var history = Get<ExplorerClosedTabHistory>(fixture.Watcher, "_closedTabs");
        history.Push(Tab(1, @"C:\unavailable"));
        history.Push(Tab(2, @"\\server\share"));
        var result = await (Task<SessionCommandResult>)fixture.Invoke("ReopenClosedTabCoreAsync", CancellationToken.None)!;
        Check.Equal(SessionCommandResult.NothingAvailable, result, "An unavailable path sends no shell command.");
        Check.That(Pop(history, out var pending) && pending.Location == @"C:\unavailable", "A failed validation does not consume history.");
        Check.That(!Pop(history, out _), "A location that can never be reopened is not offered again.");
    });

    private static Task ReopenPassesOverUnavailable() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var probed = new List<string>();
        Set(fixture.Watcher, "_recordClosedTabs", true);
        Set(fixture.Watcher, "_sessionLocationPolicy", new ExplorerSessionLocationPolicy(location =>
        {
            lock (probed) probed.Add(location);
            return location is @"C:\older" or @"C:\oldest";
        }));
        var history = Get<ExplorerClosedTabHistory>(fixture.Watcher, "_closedTabs");
        history.Push(Tab(1, @"C:\oldest"));
        history.Push(Tab(2, @"C:\older"));
        history.Push(Tab(3, @"\\server\share"));
        history.Push(Tab(4, @"C:\unplugged"));
        var candidates = history.Peek();
        var index = await (Task<int>)fixture.Invoke("FindReopenableClosedTabAsync", candidates, CancellationToken.None)!;
        Check.Equal(2, index, "The newest close that can be opened is chosen instead of failing on the newest one.");
        Check.That(Tokens(history).SequenceEqual([4, 3, 2, 1]), "Choosing does not consume or forget any close yet.");
        lock (probed)
            Check.That(probed.SequenceEqual([@"C:\unplugged", @"C:\older"]),
                "Checking stops at the chosen close, and a share is never probed.");
    });

    private static Task ReopenStopsAtUndecidedFolder() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var gate = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var exited = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var probed = new List<string>();
        var policy = new ExplorerSessionLocationPolicy(location =>
        {
            lock (probed) probed.Add(location);
            if (location != @"C:\stalled")
                return true;
            entered.TrySetResult();
            gate.Wait();
            exited.TrySetResult();
            return true;
        });
        Set(fixture.Watcher, "_recordClosedTabs", true);
        Set(fixture.Watcher, "_sessionLocationPolicy", policy);
        var history = Get<ExplorerClosedTabHistory>(fixture.Watcher, "_closedTabs");
        history.Push(Tab(1, "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"));
        history.Push(Tab(2, @"\\server\share"));
        history.Push(Tab(3, @"C:\newest"));
        try
        {
            // An earlier check is still blocked on a drive that does not answer.
            var stalled = policy.FindAvailableAsync(new ExplorerSession { Locations = [@"C:\stalled"] }, CancellationToken.None, 40);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(2));
            await stalled;
            var index = await (Task<int>)fixture.Invoke("FindReopenableClosedTabAsync", history.Peek(), CancellationToken.None)!;
            Check.Equal(-1, index, "A folder that could not be checked in time is not passed over, not even for an older This PC tab.");
            Check.That(Tokens(history).SequenceEqual([3, 1]), "Closes that can still open keep their places; only the share is forgotten.");
            lock (probed)
                Check.That(probed.SequenceEqual([@"C:\stalled"]), "Nothing is probed behind the blocked check.");
        }
        finally { gate.Set(); await exited.Task.WaitAsync(TimeSpan.FromSeconds(2)); }
    });

    private static Task SerializesRequests() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        Set(fixture.Watcher, "_captureSessions", true);
        fixture.SetCatalog();
        fixture.MarkShellConnected();
        var started = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var finish = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var command = (Func<CancellationToken, Task<SessionCommandResult>>)(async _ =>
        {
            started.TrySetResult(true);
            await finish.Task;
            throw new InvalidOperationException("injected command failure");
        });
        var first = (Task<SessionCommandResult>)fixture.Invoke("RunSessionCommandAsync", "test", command)!;
        await started.Task.WaitAsync(TimeSpan.FromSeconds(3));
        Check.Equal(SessionCommandResult.Busy, await fixture.Watcher.RestoreLastSessionAsync(), "Overlapping commands are rejected.");
        finish.TrySetResult(true);
        Check.Equal(SessionCommandResult.Failed, await first, "Command failures are not reported as success.");
        var next = (Func<CancellationToken, Task<SessionCommandResult>>)(_ => Task.FromResult(SessionCommandResult.NothingSaved));
        Check.Equal(SessionCommandResult.NothingSaved,
            await (Task<SessionCommandResult>)fixture.Invoke("RunSessionCommandAsync", "test", next)!, "The command gate is released after failure.");
        Set(fixture.Watcher, "_captureSessions", false);
        Check.Equal(SessionCommandResult.NotReady, await fixture.Watcher.ReopenClosedTabAsync(), "Stopping capture prevents late commands.");
    });

    private static Task IndependentScope() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        Set(fixture.Watcher, "_independentOpens", 1);
        Check.That((bool)fixture.Invoke("IsIndependentOpenRequested")!, "A delayed request remains protected until its scope ends.");
        return Task.CompletedTask;
    });

    private static T Get<T>(ExplorerWatcher watcher, string name) =>
        (T)typeof(ExplorerWatcher).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(watcher)!;
    private static void Set(ExplorerWatcher watcher, string name, object value) =>
        typeof(ExplorerWatcher).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(watcher, value);
}
