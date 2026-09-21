using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

internal static class ExplorerNativeFocusTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("native file-location focus activates a later existing tab without registration", () => ReusesFocusedTab(1));
        yield return ("native file-location focus repeatedly activates the third existing tab", () => ReusesFocusedTab(2));
        yield return ("native file-location item focus activates an existing tab without a client-focus event", () => ReusesItemFocus(0));
        yield return ("native file-location accessible child focus activates its existing tab", () => ReusesItemFocus(7));
        yield return ("native file-location reuse allows a busy Explorer to finish switching later tabs", SlowNativeReuse);
        yield return ("native file-location focus leaves active tabs unchanged", ActiveTabIsUnchanged);
        yield return ("native file-location focus cannot act after keyboard focus moves", () => RejectsStaleFocus(false));
        yield return ("disabled reuse does not activate native file-location focus", () => RejectsStaleFocus(true));
        yield return ("queued native file-location focus rechecks focus after the merge lock", QueuedFocusIsRechecked);
        yield return ("native file-location focus ignores delayed events from WinTab tab switching", IgnoresProgrammaticFocus);
    }

    private static Task ReusesFocusedTab(int index) => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var target = fixture.Window.TabAt(index);
        var browser = fixture.AddBrowser(out _, target);
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        var view = fixture.Window.CreateFolderView(index);
        fixture.Window.Show();
        Helper.RestoreWindowToForeground(fixture.Window.Handle);
        for (var attempt = 0; attempt < 3; attempt++)
        {
            fixture.Window.SetActive(0);
            SetFocus(view);
            await (Task)fixture.Invoke("TryActivateNativeFocusedTabAsync", view)!;
            Check.Equal(target, fixture.Window.ActiveTab, $"Native focus in an inactive file view must activate its exact tab, even without a new window or navigation. attempt={attempt} parent={fixture.Window.Handle} foreground={WinApi.GetForegroundWindow()} liveFocus={fixture.Invoke("HasNativeViewFocus", view, fixture.Window.Handle)}");
            Check.Equal(3, ExplorerWindowDiscovery.GetAllExplorerTabs(fixture.Window.Handle).Count(), "Native reuse must not create or close tabs.");
        }
    }, tabCount: 3);

    private static Task ReusesItemFocus(int child) => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var target = fixture.Window.TabAt(2);
        var browser = fixture.AddBrowser(out _, target);
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        var view = fixture.Window.CreateFolderView(2);
        fixture.Window.Show();
        fixture.Window.SetActive(0);
        Helper.RestoreWindowToForeground(fixture.Window.Handle);
        SetFocus(view);
        // If Explorer already holds keyboard focus on this hidden view, selecting the file can report
        // only the accessible item, without a second OBJID_CLIENT / CHILDID_SELF notification.
        fixture.Invoke("OnWindowShown", (nint)0, (uint)WinApi.EVENT_OBJECT_FOCUS, view, 8301, child, 0u,
            unchecked((uint)Environment.TickCount));
        var active = await Helper.DoUntilConditionAsync(() => fixture.Window.ActiveTab,
            handle => handle == target, 1_500, 20);
        Check.Equal(target, active, "A native file selection must activate its hidden tab even without a client-focus notification.");
        Check.Equal(3, ExplorerWindowDiscovery.GetAllExplorerTabs(fixture.Window.Handle).Count(), "Native reuse must not create duplicate tabs.");
    }, tabCount: 3);

    private static Task SlowNativeReuse() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var target = fixture.Window.TabAt(2);
        var browser = fixture.AddBrowser(out _, target);
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        var view = fixture.Window.CreateFolderView(2);
        fixture.Window.Show();
        fixture.Window.SetActive(0);
        Helper.RestoreWindowToForeground(fixture.Window.Handle);
        SetFocus(view);
        fixture.Window.SwitchDelayMs = 400;
        await (Task)fixture.Invoke("TryActivateNativeFocusedTabAsync", view)!;
        Check.Equal(target, fixture.Window.ActiveTab,
            $"A fresh native reuse request needs the normal tab-switch budget, not a one-second partial switch. parent={fixture.Window.Handle} foreground={WinApi.GetForegroundWindow()} liveFocus={fixture.Invoke("HasNativeViewFocus", view, fixture.Window.Handle)}");
    }, tabCount: 3);

    private static Task ActiveTabIsUnchanged() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out _, fixture.Window.FirstTab);
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        var view = fixture.Window.CreateFolderView(0);
        fixture.Window.Show();
        fixture.Window.SetActive(0);
        Helper.RestoreWindowToForeground(fixture.Window.Handle);
        SetFocus(view);
        await (Task)fixture.Invoke("TryActivateNativeFocusedTabAsync", view)!;
        Check.Equal(fixture.Window.FirstTab, fixture.Window.ActiveTab, "Ordinary file focus must not switch tabs.");
    });

    private static Task RejectsStaleFocus(bool disabled) => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out _, fixture.Window.FirstTab);
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        if (disabled) fixture.Watcher.SetReuseTabs(false);
        var view = fixture.Window.CreateFolderView(0);
        fixture.Window.Show();
        fixture.Window.SetActive(1);
        var before = fixture.Window.ActiveTab;
        Helper.RestoreWindowToForeground(fixture.Window.Handle);
        SetFocus(disabled ? view : before);
        await (Task)fixture.Invoke("TryActivateNativeFocusedTabAsync", view)!;
        Check.Equal(before, fixture.Window.ActiveTab, "Stale focus or disabled reuse must not switch tabs.");
    });

    private static Task QueuedFocusIsRechecked() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out _, fixture.Window.FirstTab);
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        var view = fixture.Window.CreateFolderView(0);
        fixture.Window.Show();
        fixture.Window.SetActive(1);
        var before = fixture.Window.ActiveTab;
        Helper.RestoreWindowToForeground(fixture.Window.Handle);
        SetFocus(view);
        var openLock = fixture.OpenLock;
        await openLock.WaitAsync();
        Task pending;
        try
        {
            pending = (Task)fixture.Invoke("TryActivateNativeFocusedTabAsync", view)!;
            await Task.Delay(40);
            SetFocus(before);
        }
        finally { openLock.Release(); }
        await pending;
        Check.Equal(before, fixture.Window.ActiveTab, "A queued focus notification must not override newer user focus.");
    });

    private static Task IgnoresProgrammaticFocus() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var browser = fixture.AddBrowser(out _, fixture.Window.FirstTab);
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        var view = fixture.Window.CreateFolderView(0);
        fixture.Window.Show();
        fixture.Window.SetActive(1);
        var before = fixture.Window.ActiveTab;
        Helper.RestoreWindowToForeground(fixture.Window.Handle);
        SetFocus(view);
        var previousEvent = unchecked((uint)(Environment.TickCount - 10));
        // Even an already-active selection establishes the event boundary of a WinTab command.
        await fixture.Watcher.SelectTabByHandle(fixture.Window.Handle, before);
        SetFocus(view);
        fixture.Invoke("OnWindowShown", (nint)0, (uint)WinApi.EVENT_OBJECT_FOCUS, view, -4, 0, 0u, previousEvent);
        await Task.Delay(400);
        Check.Equal(before, fixture.Window.ActiveTab, "Delayed focus from our own switch must not start a second switch to an intermediate tab.");
        await (Task)fixture.Invoke("TryActivateNativeFocusedTabAsync", view)!;
        Check.Equal(fixture.Window.FirstTab, fixture.Window.ActiveTab, "A fresh external request with the same live view must still be accepted.");
    });

    [DllImport("user32.dll")] private static extern nint SetFocus(nint window);
}
