using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;

internal static class NavigationNativeSelectionTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("navigation native selection activates the appended tab before its title is published", () => SelectsUntitledTab(3));
        yield return ("navigation native selection works beyond the ninth tab", () => SelectsUntitledTab(12));
        yield return ("navigation middle-click activates the tab of a window that is not yet in front", SelectsInWindowNotYetInFront);
        yield return ("navigation expired accessibility work cannot block later clicks", ExpiredWorkCannotBlock);
        yield return ("navigation middle-click accepts any control inside the active tab and nothing outside it", AcceptsWholeActiveTab);
    }

    private static async Task AcceptsWholeActiveTab()
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(() =>
        {
            using var window = new ExplorerTabActivationTests.ActivationWindow(2);
            window.SetActive(0);
            var activeView = window.CreateFolderView(0);
            var backgroundView = window.CreateFolderView(1);
            Check.That(ExplorerNavigationAccess.IsInsideActiveTab(activeView, window.Handle),
                "A folder view nested deep inside the active tab is a valid middle-click target, not only the navigation tree.");
            Check.That(ExplorerNavigationAccess.IsInsideActiveTab(window.FirstTab, window.Handle),
                "The active tab window itself counts as inside the active tab.");
            Check.That(!ExplorerNavigationAccess.IsInsideActiveTab(backgroundView, window.Handle),
                "A control of a background tab cannot be under the pointer of the active tab.");
            Check.That(!ExplorerNavigationAccess.IsInsideActiveTab(window.Handle, window.Handle),
                "The frame (tab strip, title bar) belongs to no tab, so a tab-closing middle click is ignored.");
            Check.That(!ExplorerNavigationAccess.IsInsideActiveTab(0, window.Handle) && !ExplorerNavigationAccess.IsInsideActiveTab(activeView, 0),
                "Missing handles are never inside a tab.");
        }, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }

    /// <summary>
    /// A middle click on an Explorer window behind another application makes Explorer come to the front only
    /// after the button is released. Until then the click is still the user's, so its tab must be activated.
    /// </summary>
    private static async Task SelectsInWindowNotYetInFront()
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            using var window = new ExplorerTabActivationTests.ActivationWindow(3);
            window.Show();
            window.SetActive(0);
            Check.That(ExplorerNavigationAccess.ForegroundFrame() != window.Handle, "The window must be behind another window for this test.");
            var before = new[] { window.TabAt(0), window.TabAt(1) };
            var click = new NavigationClick(WindowIdentity.Capture(window.Handle), WindowIdentity.Capture(window.FirstTab),
                WindowIdentity.Capture(window.Handle), before, default);
            Check.That(click.IsCurrent(), "A click on a window that is not yet in front is still current.");
            var result = await NavigationTabActivation.RunAsync(before, window.FirstTab, click.IsCurrent,
                () => ExplorerNavigationAccess.Observe(window.Handle),
                tab => ExplorerNavigationAccess.SelectNewTab(click, tab, click.IsCurrent), CancellationToken.None, timeoutMs: 800);
            Check.Equal(NavigationActivationResult.Activated, result, "The click's new tab is activated although its window is not in front yet.");
            Check.Equal(window.TabAt(2), window.ActiveTab, "The appended tab must be selected.");
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }

    private static async Task SelectsUntitledTab(int tabCount)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            using var window = new ExplorerTabActivationTests.ActivationWindow(tabCount);
            window.Show();
            window.SetActive(0);
            Helper.RestoreWindowToForeground(window.Handle);
            var before = Enumerable.Range(0, tabCount - 1).Select(window.TabAt).ToArray();
            var click = new NavigationClick(WindowIdentity.Capture(window.Handle), WindowIdentity.Capture(window.FirstTab),
                WindowIdentity.Capture(window.Handle), before, default);
            var result = await NavigationTabActivation.RunAsync(before, window.FirstTab, () => true,
                () => ExplorerNavigationAccess.Observe(window.Handle),
                tab => ExplorerNavigationAccess.SelectNewTab(click, tab, () => true), CancellationToken.None, timeoutMs: 800);
            Check.Equal(NavigationActivationResult.Activated, result,
                "The new native tab is known even when Explorer has not published a UIA title; do not wait for a title or scan the strip.");
            Check.Equal(window.TabAt(tabCount - 1), window.ActiveTab, "The appended tab, not a same-name existing tab, must be selected.");
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }

    private static Task ExpiredWorkCannotBlock()
    {
        using var gate = new NavigationClickGate();
        var stalled = gate.Begin(100, 0, out _)!;
        var next = gate.Begin(100, NavigationClickGate.RequestLifetimeMs + 1, out _);
        Check.That(next != null, "An expired accessibility worker must not require an extra click or its eventual completion to release the gate.");
        gate.Complete(stalled);
        Check.That(gate.IsCurrent(next!, NavigationClickGate.RequestLifetimeMs + 2), "The old worker cannot retire the new request when it finally returns.");
        gate.Complete(next!);
        return Task.CompletedTask;
    }
}
