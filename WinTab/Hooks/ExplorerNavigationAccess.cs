using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using WinTab.Helpers;
using WinTab.Interop;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// What was true at the moment the middle button went down inside a tab. <paramref name="Target"/> is the
/// control under the pointer; <paramref name="OnNavigationTree"/> tells whether it is the navigation pane's
/// tree, the one place where a click beside the folder items can be recognised before Explorer reacts.
/// </summary>
internal sealed record NavigationClick(WindowIdentity Window, WindowIdentity SourceTab, WindowIdentity Target, nint[] TabsBefore, Point Point,
    bool OnNavigationTree = false)
{
    /// <summary>
    /// The click can still be acted on: the same window, tab and target exist. The window need not be in front:
    /// a click on an Explorer window behind another application brings it to the front only after the button
    /// is released. Moving on to another window ends the click through the foreground event instead.
    /// </summary>
    public bool IsCurrent() =>
        Window.IsCurrent && SourceTab.IsCurrent && Target.IsCurrent &&
        WinApi.GetParent(SourceTab.Handle) == Window.Handle;
}

/// <summary>Reads native tab handles and navigation items. Accessibility work never runs on the hook thread.</summary>
internal static class ExplorerNavigationAccess
{
    internal const int MaxTabs = 128;
    private const uint ObjIdClient = 0xFFFFFFFC;
    private const int OutlineItemRole = 0x24;
    private const int UnavailableOrInvisible = 0x1 | 0x8000 | 0x10000;

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
    /// Whether the point is on a folder item of the navigation pane, where a middle click makes Explorer
    /// open a tab. The item's label and its icon count; the expansion glyph beside them does not.
    /// </summary>
    public static bool IsFolderItemAt(nint tree, Point point)
    {
        var accessible = AccessibleTree(tree);
        try
        {
            var child = accessible.accHitTest(point.X, point.Y);
            try
            {
                if (child is not int id)
                    return false;
                var item = ReadItem(accessible, id, tree);
                return item != null && item.HitBounds.Contains(point);
            }
            finally
            {
                if (child != null && Marshal.IsComObject(child))
                    Marshal.ReleaseComObject(child);
            }
        }
        finally { Marshal.ReleaseComObject(accessible); }
    }

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
        // Explorer may still be moving the new tab's window; that is no evidence either way.
        if (ReadTabs(window) is not { } tabs)
            return NavigationSelectOutcome.NotReady;
        if (!SameHandles(tabs, expected) || tabs[0] != click.SourceTab.Handle)
            return NavigationSelectOutcome.Rejected;
        return WinApi.PostMessage(window, WinApi.WM_COMMAND, 0xA221, expected.Length)
            ? NavigationSelectOutcome.Selected
            : NavigationSelectOutcome.Rejected;
    }

    private sealed record NavigationItem(int Id, Rectangle HitBounds);

    private static NavigationItem? ReadItem(IAccessible accessible, int id, nint tree)
    {
        if (id <= 0 || Convert.ToInt32(accessible.get_accRole(id)) != OutlineItemRole ||
            (Convert.ToInt32(accessible.get_accState(id)) & UnavailableOrInvisible) != 0) return null;
        accessible.accLocation(out var left, out var top, out var width, out var height, id);
        if (width <= 0 || height <= 0 || !WinApi.GetWindowRect(tree, out var bounds)) return null;
        var dpi = Math.Max(96u, WinApi.GetDpiForWindow(tree));
        // MSAA gives the label rectangle. Include its adjacent small icon, not the expansion glyph to its left.
        var iconWidth = GetSystemMetricsForDpi(WinApi.SM_CXSMICON, dpi) + (int)(4 * dpi / 96);
        var hit = Rectangle.Intersect(new Rectangle(left - iconWidth, top, width + iconWidth, height),
            new Rectangle(bounds.Left, bounds.Top, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top));
        return hit.IsEmpty ? null : new NavigationItem(id, hit);
    }

    private static IAccessible AccessibleTree(nint tree)
    {
        var guid = typeof(IAccessible).GUID;
        Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(tree, ObjIdClient, ref guid, out var accessible));
        return accessible;
    }

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(nint window, uint objectId, ref Guid interfaceId,
        [MarshalAs(UnmanagedType.Interface)] out IAccessible accessible);
    [DllImport("user32.dll")]
    private static extern int GetSystemMetricsForDpi(int index, uint dpi);
}
