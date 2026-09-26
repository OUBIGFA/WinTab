using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// What was true at the moment the middle button went down inside a tab. <paramref name="Target"/> is the
/// control under the pointer; <paramref name="OnNavigationTree"/> tells whether it is the navigation pane's tree.
/// </summary>
internal sealed record NavigationClick(WindowIdentity Window, WindowIdentity SourceTab, WindowIdentity Target, nint[] TabsBefore, Point Point,
    bool OnNavigationTree = false)
{
    /// <summary>
    /// The click can still be acted on: the same window and tab exist. The control that was clicked need not:
    /// Explorer may rebuild an address-bar or Home-page control while it opens the tab, and the tab is still the
    /// click's. The window need not be in front either: a click on an Explorer window behind another application
    /// brings it to the front only after the button is released. Moving on to another window ends the click
    /// through the foreground event instead.
    /// </summary>
    public bool IsCurrent() =>
        Window.IsCurrent && SourceTab.IsCurrent && WinApi.GetParent(SourceTab.Handle) == Window.Handle;
}

/// <summary>Reads native tab handles and selects the tab a click opened. Only window handles are read on the input thread.</summary>
internal static class ExplorerNavigationAccess
{
    internal const int MaxTabs = 128;

    public static nint ActiveTab(nint window) => WinApi.FindWindowEx(window, 0, "ShellTabWindowClass", null);

    public static nint ForegroundFrame()
    {
        var foreground = WinApi.GetForegroundWindow();
        var root = WinApi.GetAncestor(foreground, WinApi.GA_ROOT);
        return root != 0 ? root : foreground;
    }

    /// <summary>The window's tabs as one snapshot, the active tab first; null while Explorer is moving tab windows.</summary>
    public static nint[]? ReadTabs(nint window) => ExplorerWindowDiscovery.GetStableExplorerTabs(window, MaxTabs + 1);

    public static bool SameHandles(nint[] first, nint[] second) => first.Length == second.Length && new HashSet<nint>(first).SetEquals(second);

    /// <summary>Whether <paramref name="tree"/> is the navigation pane's tree control of <paramref name="window"/>.</summary>
    public static bool IsNavigationTree(nint tree, nint window) =>
        tree != 0 && window != 0 &&
        WinApi.IsWindowHasClassName(tree, "SysTreeView32") &&
        WinApi.IsWindowHasClassName(WinApi.GetParent(tree), "NamespaceTreeControl") &&
        WinApi.GetAncestor(tree, WinApi.GA_ROOT) == window;

    /// <summary>
    /// Whether <paramref name="hit"/> lies inside the active tab of <paramref name="window"/>: the navigation
    /// pane, the file list, the address bar or the Home page. The tab strip and the frame belong to no tab,
    /// so a middle click that closes a tab is never mistaken for one that opens a folder.
    /// </summary>
    public static bool IsInsideActiveTab(nint hit, nint window)
    {
        if (hit == 0 || window == 0)
            return false;
        var activeTab = ActiveTab(window);
        if (activeTab == 0 || WinApi.GetParent(activeTab) != window)
            return false;
        var current = hit;
        for (var depth = 0; current != 0 && current != window && depth < 64; depth++)
        {
            if (current == activeTab)
                return true;
            current = WinApi.GetParent(current);
        }
        return false;
    }

    /// <summary>
    /// Whether <paramref name="hit"/> is a control in the Explorer frame that can be observed for a native
    /// middle-click navigation. Windows 11's address bar is hosted by the frame/XAML island rather than below
    /// the active ShellTabWindowClass, so it cannot be rejected solely because it is not an active-tab child.
    /// The caller still requires a unique appended tab before selecting anything; this method only permits
    /// observation and never treats a frame click as a navigation by itself.
    /// </summary>
    public static bool IsPotentialNavigationTarget(nint hit, nint window)
    {
        if (hit == 0 || window == 0 || hit == window ||
            !ExplorerWindowDiscovery.IsFileExplorerWindow(window))
            return false;

        nint current = hit;
        for (var depth = 0; current != 0 && depth < 64; depth++)
        {
            if (current == window)
                return true;
            current = WinApi.GetParent(current);
        }
        return false;
    }

    /// <summary>The tabs and the active tab from one snapshot; null when no consistent snapshot could be read.</summary>
    public static NavigationTabObservation? Observe(nint window) =>
        ReadTabs(window) is { } tabs ? new(tabs, tabs.Length > 0 ? tabs[0] : 0) : null;

    /// <summary>
    /// Explorer appends a navigation middle-click tab. After its unique native handle is confirmed, select
    /// the appended index with Explorer's own command; no UIA tree, title publication or COM call is needed.
    /// The caller still verifies the exact active HWND after the command, rather than claiming success here.
    /// </summary>
    public static NavigationSelectOutcome SelectNewTab(NavigationClick click, nint newTab, Func<bool> isCurrent)
    {
        var window = click.Window.Handle;
        nint[] expected = [.. click.TabsBefore, newTab];
        if (!isCurrent() || WinApi.GetParent(newTab) != window)
            return NavigationSelectOutcome.Rejected;
        // The native tab HWND appears before its view is ready. Do not select a half-created view, even
        // if its UI thread happens to answer sent messages while it is still initializing the tab.
        if (!WinApi.IsWindowVisible(newTab))
            return NavigationSelectOutcome.NotReady;
        // Do not park a switch in an already busy Explorer's queue: the click may be cancelled before it
        // reaches that command. Probe without side effects, then recheck the lease and tabs before posting.
        if (!WinApi.TrySendMessage(window, WinApi.WM_NULL, 0, 0, 20))
            return isCurrent() ? NavigationSelectOutcome.NotReady : NavigationSelectOutcome.Rejected;
        if (!isCurrent())
            return NavigationSelectOutcome.Rejected;
        // Explorer may still be moving the new tab's window; that is no evidence either way.
        if (ReadTabs(window) is not { } tabs)
            return NavigationSelectOutcome.NotReady;
        if (!SameHandles(tabs, expected) || tabs[0] != click.SourceTab.Handle || !isCurrent())
            return NavigationSelectOutcome.Rejected;
        // Explorer must dispatch this command in its normal message loop. A synchronous send can reenter
        // tab initialization (even after WM_NULL answered) and leave the new view permanently uninitialized.
        return WinApi.PostMessage(window, WinApi.WM_COMMAND, 0xA221, expected.Length)
            ? NavigationSelectOutcome.Selected
            : NavigationSelectOutcome.Rejected;
    }
}
