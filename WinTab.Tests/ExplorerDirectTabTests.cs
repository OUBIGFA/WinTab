using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using Fixture = ExplorerTabLifetimeTests.Fixture;

/// <summary>
/// A merged folder is opened as a tab that Explorer creates at the folder itself, appended in the background,
/// and brought to the front only once it is there; the window never shows a default page in between. Where
/// Explorer does not understand the request it answers with a window or with nothing, and the classic new-tab
/// command takes over for good.
/// </summary>
internal static class ExplorerDirectTabTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("an unrelated preloaded frame cannot complete an accepted direct request", PreloadedFrameDoesNotAnswerRequest);
        yield return ("a tab Explorer appends for a direct request is the one returned", AppendedTabIsReturned);
        yield return ("direct tab creation refuses two concurrent additions", () => AmbiguousTabsAreNotClaimed(false, false));
        yield return ("classic tab creation refuses two concurrent additions", () => AmbiguousTabsAreNotClaimed(true, false));
        yield return ("direct tab creation refuses a replaced old tab", () => AmbiguousTabsAreNotClaimed(false, true));
        yield return ("classic tab creation refuses a replaced old tab", () => AmbiguousTabsAreNotClaimed(true, true));
        yield return ("classic tab creation returns the unique addition", ClassicUniqueTabIsReturned);
        yield return ("a direct tab at another location is never redirected", DirectTabAtAnotherLocationIsNotNavigated);
        yield return ("a window instead of a tab retires direct tab creation", WindowInsteadOfTabRetiresDirectTabs);
        yield return ("no answer to a direct tab request retires direct tab creation", NoAnswerRetiresDirectTabs);
        yield return ("a retired direct request is not issued again", RetiredRequestIsNotIssued);
        yield return ("a window without a tracked tab gets the classic command without retiring direct tabs", UntrackedWindowKeepsDirectTabs);
        yield return ("an appended tab is activated by its index with a single command", AppendedTabIsActivatedByIndex);
        yield return ("a slow Explorer tab switch still activates the appended tab", SlowSwitchStillActivates);
        yield return ("an appended tab whose index moved is still activated", MovedIndexStillActivates);
        yield return ("an appended index fallback selects before foreground activation", FallbackSelectsBeforeForeground);
        yield return ("tab creation foregrounds the settled frame before any new tab exists", CreationPreparesSettledFrame);
        yield return ("a tab created at its location is not navigated a second time", TabCreatedAtLocationIsNotNavigatedAgain);
        yield return ("a tab created elsewhere is navigated to the location", TabCreatedElsewhereIsNavigated);
        yield return ("a failed tab is still closed after the merge's own deadline has passed", FailedTabIsClosedAfterMergeDeadline);
        yield return ("failed tab cleanup stops when the hook is stopped", FailedTabCleanupHonoursHookStop);
        yield return ("failed tab cleanup leaves a tab that has changed identity alone", FailedTabCleanupChecksIdentity);
    }

    private static Task PreloadedFrameDoesNotAnswerRequest() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var window = fixture.Window;
        var knownTabs = new HashSet<nint>(ExplorerWindowDiscovery.GetAllExplorerTabs(window.Handle));
        var windows = new List<nint> { window.Handle };
        fixture.SetExplorerWindowSource(() => windows.ToArray());
        var wait = (Task<nint>)fixture.Invoke("WaitForTabAtLocationAsync", window.Handle, knownTabs,
            new HashSet<nint>(windows), Fixture.Location, 2_000)!;
        using var preloaded = new RemoteExplorerFrame(visible: false);
        windows.Add(preloaded.Handle);
        await Task.Delay(500);
        var appended = window.AddTab();
        Check.Equal(appended, await wait, "An unrelated hidden frame must not cause a second new-tab command before the real tab arrives.");
        Check.That(!fixture.DirectTabUnsupported, "Preloading does not mean direct tabs are unsupported.");
    }, tabCount: 1);

    private static Task AppendedTabIsReturned() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var window = fixture.Window;
        var knownTabs = new HashSet<nint>(ExplorerWindowDiscovery.GetAllExplorerTabs(window.Handle));
        fixture.SetExplorerWindowSource(() => [window.Handle]);
        var knownWindows = new HashSet<nint> { window.Handle };

        var wait = (Task<nint>)fixture.Invoke("WaitForTabAtLocationAsync", window.Handle, knownTabs, knownWindows, Fixture.Location, 2_000)!;
        await Task.Delay(150);
        Check.That(!wait.IsCompleted, "The wait must not finish before Explorer has appended a tab.");
        var appended = window.AddTab();

        Check.Equal(appended, await wait, "The appended tab must be the one returned.");
        Check.That(!fixture.DirectTabUnsupported, "A successful request must keep direct tab creation enabled.");
        Check.Equal(0, fixture.Statuses.Count, "A successful request must not report anything: " + string.Join(" | ", fixture.Statuses));
    }, tabCount: 1);

    private static Task AmbiguousTabsAreNotClaimed(bool classic, bool replaceOld) => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var window = fixture.Window;
        var knownTabs = new HashSet<nint>(ExplorerWindowDiscovery.GetAllExplorerTabs(window.Handle));
        fixture.SetExplorerWindowSource(() => [window.Handle]);
        var wait = classic
            ? ExplorerWindowDiscovery.ListenForNewExplorerTabAsync(window.Handle, knownTabs, 500)
            : (Task<nint>)fixture.Invoke("WaitForTabAtLocationAsync", window.Handle, knownTabs,
                new HashSet<nint> { window.Handle }, Fixture.Location, 500)!;
        var first = window.AddTab();
        var second = window.AddTab();
        if (replaceOld)
            Check.That(WinTab.WinAPI.WinApi.TrySendMessage(window.FirstTab, WinTab.WinAPI.WinApi.WM_CLOSE, 0, 0),
                "The old tab must close before the next observation.");
        var rejected = false;
        try { await wait; }
        catch (InvalidOperationException) { rejected = true; }
        Check.That(rejected, "Competing tab changes must abort the merge rather than claim a user's tab.");
        var remaining = ExplorerWindowDiscovery.GetAllExplorerTabs(window.Handle).ToArray();
        Check.That(remaining.Contains(first) && remaining.Contains(second), "Neither unowned new tab may be closed.");
        Check.That(!fixture.DirectTabUnsupported, "A competing user action must not disable direct tab support or request a fallback tab.");
    }, tabCount: 1);

    private static Task ClassicUniqueTabIsReturned() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var window = fixture.Window;
        var wait = ExplorerWindowDiscovery.ListenForNewExplorerTabAsync(window.Handle,
            ExplorerWindowDiscovery.GetAllExplorerTabs(window.Handle).ToArray(), 500);
        var appended = window.AddTab();
        Check.Equal(appended, await wait, "A unique new tab must still be claimed on the classic path.");
    }, tabCount: 1);

    private static Task DirectTabAtAnotherLocationIsNotNavigated() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var navigations = 0;
        var browser = fixture.CreateBrowser((method, _) => method switch
        {
            "get_HWND" => (long)fixture.Window.Handle,
            "get_LocationURL" => "file:///C:/Users",
            "get_Document" => null,
            "Navigate2" => Count(ref navigations),
            "add_NavigateComplete2" or "remove_NavigateComplete2" => null,
            _ => throw new InvalidOperationException("Unexpected browser call: " + method)
        });
        Check.That(!await (Task<bool>)fixture.Invoke("NavigateNewTabToTargetAsync", browser, Fixture.Location, true)!,
            "A direct tab that never reaches the requested location must not count as the merge's result.");
        Check.Equal(0, navigations, "An unconfirmed direct tab may belong to the user and must never be redirected.");
    });

    private static Task WindowInsteadOfTabRetiresDirectTabs() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var window = fixture.Window;
        var knownTabs = new HashSet<nint>(ExplorerWindowDiscovery.GetAllExplorerTabs(window.Handle));
        var windows = new List<nint> { window.Handle };
        fixture.SetExplorerWindowSource(() => windows.ToArray());
        var knownWindows = new HashSet<nint>(windows);

        var wait = (Task<nint>)fixture.Invoke("WaitForTabAtLocationAsync", window.Handle, knownTabs, knownWindows, Fixture.Location, 2_000)!;
        await Task.Delay(100);
        // Explorer's answer is a window it has shown; a hidden frame would be a preload, not an answer.
        using var opened = new RemoteExplorerFrame(visible: true);
        windows.Add(opened.Handle);

        Check.Equal<nint>(0, await wait, "A window is not the requested tab.");
        Check.That(fixture.DirectTabUnsupported, "Explorer answering with a window must retire direct tab creation.");
        Check.That(fixture.Statuses.Any(status => status.Contains("window", StringComparison.OrdinalIgnoreCase)),
            "The fallback must be reported: " + string.Join(" | ", fixture.Statuses));
    }, tabCount: 1);

    private static Task NoAnswerRetiresDirectTabs() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var window = fixture.Window;
        var knownTabs = new HashSet<nint>(ExplorerWindowDiscovery.GetAllExplorerTabs(window.Handle));
        fixture.SetExplorerWindowSource(() => [window.Handle]);
        var knownWindows = new HashSet<nint> { window.Handle };

        var tab = await (Task<nint>)fixture.Invoke("WaitForTabAtLocationAsync", window.Handle, knownTabs, knownWindows, Fixture.Location, 300)!;

        Check.Equal<nint>(0, tab, "Without a tab nothing can be returned.");
        Check.That(fixture.DirectTabUnsupported, "Explorer answering with nothing must retire direct tab creation.");
        Check.Equal(1, fixture.Statuses.Count, "The fallback must be reported once: " + string.Join(" | ", fixture.Statuses));
    }, tabCount: 1);

    private static Task RetiredRequestIsNotIssued() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var window = fixture.Window;
        fixture.AddBrowser(out _, window.FirstTab);
        fixture.DirectTabUnsupported = true;
        var tabs = ExplorerWindowDiscovery.GetAllExplorerTabs(window.Handle).ToArray();

        var tab = await (Task<nint>)fixture.Invoke("CreateTabAtLocationAsync", window.Handle, WindowIdentity.Capture(window.Handle), tabs, Fixture.Location)!;

        Check.Equal<nint>(0, tab, "A retired request must hand over to the classic command at once.");
        Check.Equal(0, fixture.Statuses.Count, "Nothing new must be reported: " + string.Join(" | ", fixture.Statuses));
    }, tabCount: 1);

    private static Task UntrackedWindowKeepsDirectTabs() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var window = fixture.Window;
        var tabs = ExplorerWindowDiscovery.GetAllExplorerTabs(window.Handle).ToArray();

        var tab = await (Task<nint>)fixture.Invoke("CreateTabAtLocationAsync", window.Handle, WindowIdentity.Capture(window.Handle), tabs, Fixture.Location)!;

        Check.Equal<nint>(0, tab, "Without a tracked tab there is no browser to ask.");
        Check.That(!fixture.DirectTabUnsupported, "A window that is not tracked yet says nothing about Explorer's support.");
        Check.Equal(0, fixture.Statuses.Count, "Nothing must be reported: " + string.Join(" | ", fixture.Statuses));
    }, tabCount: 1);

    private static Task AppendedTabIsActivatedByIndex() => WithFrame(async (fixture, frame) =>
    {
        var target = frame.TabAt(2);
        frame.SetActive(0);

        var activated = await Activate(fixture, frame, target, 2);

        Check.That(activated, $"The appended tab must be activated; commands={string.Join(',', frame.SwitchCommands)} trace={frame.Trace}");
        Check.Equal(target, frame.ActiveTab, "The appended tab must be the active one.");
        Check.That(frame.SwitchCommands.SequenceEqual([2]), $"Only the appended index may be requested; commands={string.Join(',', frame.SwitchCommands)}");
    });

    private static Task SlowSwitchStillActivates() => WithFrame(async (fixture, frame) =>
    {
        var target = frame.TabAt(2);
        frame.SetActive(0);
        frame.SwitchDelayMs = 450;

        var activated = await Activate(fixture, frame, target, 2);

        Check.That(activated, $"A slow switch must still succeed; commands={string.Join(',', frame.SwitchCommands)} trace={frame.Trace}");
        Check.Equal(target, frame.ActiveTab, "The appended tab must be active after the slow switch.");
        Check.That(frame.SwitchCommands.SequenceEqual([2]), $"A slow switch must not be repeated; commands={string.Join(',', frame.SwitchCommands)}");
    });

    private static Task MovedIndexStillActivates() => WithFrame(async (fixture, frame) =>
    {
        var target = frame.TabAt(2);
        frame.SetActive(0);

        // The index describes another tab by now; the outcome is verified and the tab found by cycling.
        var activated = await Activate(fixture, frame, target, 1);

        Check.That(activated, $"A moved tab must still be activated; commands={string.Join(',', frame.SwitchCommands)} trace={frame.Trace}");
        Check.Equal(target, frame.ActiveTab, "The appended tab must end up active even when its index moved.");
    });

    private static Task FallbackSelectsBeforeForeground() => WithFrame(async (fixture, frame) =>
    {
        var target = frame.TabAt(2);
        frame.SetActive(0);
        fixture.Window.Show();
        Helper.RestoreWindowToForeground(fixture.Window.Handle);
        frame.TabsOnActivation.Clear();
        Check.That(await Activate(fixture, frame, target, 1), "An outdated appended index must still resolve the correct tab.");
        Check.That(frame.TabsOnActivation.TryPeek(out var firstActivated) && firstActivated == target,
            "The half-created background tab must be selected before foregrounding, including the index fallback.");
        Check.Equal(1, frame.TabsOnActivation.Count, "Selection recovery must not activate the frame twice.");
    });

    private static Task CreationPreparesSettledFrame() => WithFrame(async (fixture, frame) =>
    {
        fixture.Window.Show();
        Helper.RestoreWindowToForeground(fixture.Window.Handle);
        frame.SetActive(0);
        frame.TabsOnActivation.Clear();
        var identity = WindowIdentity.Capture(frame.Handle);
        Check.That(await (Task<bool>)fixture.Invoke("PrepareWindowForTabCreationAsync", identity)!,
            "The settled frame must reach foreground before BrowseObject is allowed to create a tab.");
        Check.Equal(frame.Handle, WinTab.WinAPI.WinApi.GetForegroundWindow());
        Check.Equal(0, frame.SwitchCommands.Count, "Preparation must not cycle any existing tabs.");
        Check.Equal(frame.TabAt(0), frame.ActiveTab);
        Check.That(await (Task<bool>)fixture.Invoke("PrepareWindowForTabCreationAsync", identity)!, "An already foreground frame stays ready.");
        Check.Equal(1, frame.TabsOnActivation.Count, "An already foreground frame must not be activated again.");
    });

    private static Task TabCreatedAtLocationIsNotNavigatedAgain() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var reads = 0;
        var browser = fixture.CreateBrowser((method, _) => method switch
        {
            "get_HWND" => (long)fixture.Window.Handle,
            // Explorer is still on its way to the folder for the first reads.
            "get_LocationURL" => ++reads < 4 ? string.Empty : "file:///C:/WinTab-lifetime",
            "get_Document" => null,
            "Navigate2" => throw new InvalidOperationException("A tab created at its location must not be navigated again."),
            _ => throw new InvalidOperationException("Unexpected browser call: " + method)
        });

        var navigated = await (Task<bool>)fixture.Invoke("NavigateNewTabToTargetAsync", browser, Fixture.Location, true)!;

        Check.That(navigated, "Arriving at the location on its own must count as navigated.");
        Check.That(reads >= 4, "The location must have been observed until it arrived.");
    });

    private static Task TabCreatedElsewhereIsNavigated() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var navigations = 0;
        var browser = fixture.CreateBrowser((method, _) => method switch
        {
            "get_HWND" => (long)fixture.Window.Handle,
            "get_LocationURL" => navigations == 0 ? string.Empty : "file:///C:/WinTab-lifetime",
            "get_Document" => null,
            "Navigate2" => Count(ref navigations),
            "add_NavigateComplete2" or "remove_NavigateComplete2" => null,
            _ => throw new InvalidOperationException("Unexpected browser call: " + method)
        });

        var navigated = await (Task<bool>)fixture.Invoke("NavigateNewTabToTargetAsync", browser, Fixture.Location, false)!;

        Check.That(navigated, "A tab that is not at the location must be navigated there.");
        Check.Equal(1, navigations, "The tab must be navigated exactly once.");
    });

    /// <summary>
    /// Runs the cleanup inside a merge whose deadline has already passed. The tab is the merge's own, so it is
    /// closed anyway; the merge's cancellation must not be what decides whether the tab is left behind.
    /// </summary>
    private static Task FailedTabIsClosedAfterMergeDeadline() => WithFrame(async (fixture, frame) =>
    {
        var tab = frame.TabAt(2);
        var tabIdentity = WindowIdentity.Capture(tab);
        using var expired = new MergeOperation(WindowIdentity.Capture(frame.Handle), 0, CancellationToken.None, () => true, 1);
        await Task.Delay(30);
        Check.That(expired.Token.IsCancellationRequested, "The merge must have run out of time before the cleanup starts.");
        var context = (AsyncLocal<MergeOperation?>)typeof(ExplorerWatcher)
            .GetField("_currentMerge", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.GetValue(fixture.Watcher)!;
        context.Value = expired;
        try
        {
            await (Task)fixture.Invoke("CloseFailedNewTabAsync", WindowIdentity.Capture(frame.Handle), tabIdentity)!;
        }
        finally
        {
            context.Value = null;
        }

        Check.Equal(1, frame.CloseTabCommandCount, "The failed tab must receive exactly one close command.");
        Check.That(!tabIdentity.IsCurrent, "The failed tab must be gone once the cleanup returns.");
        Check.Equal(2, ExplorerWindowDiscovery.GetAllExplorerTabs(frame.Handle).Count(), "Only the failed tab may be closed.");
        Check.Equal(0, fixture.Statuses.Count, "A completed cleanup has nothing to report: " + string.Join(" | ", fixture.Statuses));
    });

    private static Task FailedTabCleanupHonoursHookStop() => WithFrame(async (fixture, frame) =>
    {
        var tab = frame.TabAt(2);
        var tabIdentity = WindowIdentity.Capture(tab);
        fixture.HookLifetime.Cancel();

        await (Task)fixture.Invoke("CloseFailedNewTabAsync", WindowIdentity.Capture(frame.Handle), tabIdentity)!;

        Check.Equal(0, frame.CloseTabCommandCount, "No close command may be sent once the hook has been stopped.");
        Check.That(tabIdentity.IsCurrent, "The tab must be left alone when the user has stopped WinTab.");
    });

    private static Task FailedTabCleanupChecksIdentity() => WithFrame(async (fixture, frame) =>
    {
        var tabIdentity = WindowIdentity.Capture(frame.TabAt(2));
        tabIdentity.Release();

        await (Task)fixture.Invoke("CloseFailedNewTabAsync", WindowIdentity.Capture(frame.Handle), tabIdentity)!;

        Check.Equal(0, frame.CloseTabCommandCount, "A tab whose identity has been retired is not ours to close.");
        Check.Equal(3, ExplorerWindowDiscovery.GetAllExplorerTabs(frame.Handle).Count(), "All tabs must remain.");
    });

    private static object? Count(ref int counter)
    {
        counter++;
        return null;
    }

    private static Task<bool> Activate(Fixture fixture, RemoteExplorerFrame frame, nint tab, int index) =>
        (Task<bool>)fixture.Invoke("ActivateAppendedTabAsync", frame.Handle, WindowIdentity.Capture(frame.Handle),
            tab, WindowIdentity.Capture(tab), index)!;

    private static Task WithFrame(Func<Fixture, RemoteExplorerFrame, Task> body) => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true, tabCount: 3);
        await body(fixture, frame);
    });
}
