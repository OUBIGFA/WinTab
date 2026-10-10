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
        yield return ("navigation expired accessibility work cannot block later clicks", ExpiredWorkCannotBlock);
        yield return ("navigation middle-click accepts any control inside the active tab and nothing outside it", AcceptsWholeActiveTab);
    }

    private sealed class NoInput : IDisposable
    {
        public void Dispose() { }
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

    private static Task ExpiredWorkCannotBlock()
    {
        using var gate = new NavigationClickGate();
        var stalled = gate.Begin(100, 0, out _)!;
        var next = gate.Begin(100, NavigationClickGate.HoldLimitMs + 1, out _);
        Check.That(next != null, "An expired accessibility worker must not require an extra click or its eventual completion to release the gate.");
        gate.Complete(stalled);
        Check.That(gate.IsCurrent(next!, NavigationClickGate.HoldLimitMs + 2), "The old worker cannot retire the new request when it finally returns.");
        gate.Complete(next!);
        return Task.CompletedTask;
    }
}
