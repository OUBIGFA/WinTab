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
        yield return ("navigation expired accessibility work cannot block later clicks", ExpiredWorkCannotBlock);
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
        var stalled = gate.Begin(100, default, 0)!;
        var next = gate.Begin(100, default, NavigationClickGate.RequestLifetimeMs + 1);
        Check.That(next != null, "An expired accessibility worker must not require an extra click or its eventual completion to release the gate.");
        gate.Complete(stalled, false);
        Check.That(gate.IsCurrent(next!, NavigationClickGate.RequestLifetimeMs + 2), "The old worker cannot retire the new request when it finally returns.");
        gate.Complete(next!, true);
        return Task.CompletedTask;
    }
}
