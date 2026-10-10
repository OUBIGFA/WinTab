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
        yield return ("no answer to a direct tab request retires direct tab creation", NoAnswerRetiresDirectTabs);
        yield return ("a retired direct request is not issued again", RetiredRequestIsNotIssued);
        yield return ("a window without a tracked tab gets the classic command without retiring direct tabs", UntrackedWindowKeepsDirectTabs);
        yield return ("a tab created at its location is not navigated a second time", TabCreatedAtLocationIsNotNavigatedAgain);
        yield return ("a tab created elsewhere is navigated to the location", TabCreatedElsewhereIsNavigated);
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

    private static object? Count(ref int counter)
    {
        counter++;
        return null;
    }

    private static Task<bool> Activate(Fixture fixture, RemoteExplorerFrame frame, nint tab, int index) =>
        (Task<bool>)fixture.Invoke("ActivateAppendedTabAsync", frame.Handle, WindowIdentity.Capture(frame.Handle),
            tab, WindowIdentity.Capture(tab), index)!;
}
