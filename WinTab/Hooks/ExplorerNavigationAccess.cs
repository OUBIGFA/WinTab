using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using WinTab.Helpers;
using WinTab.Interop;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>What was true at the moment the middle button went down on the navigation pane.</summary>
internal sealed record NavigationClick(WindowIdentity Window, WindowIdentity SourceTab, WindowIdentity Tree, nint[] TabsBefore, Point Point)
{
    /// <summary>The click can still be acted on: the same window, tab and tree exist and the window is in front.</summary>
    public bool IsCurrent() =>
        Window.IsCurrent && SourceTab.IsCurrent && Tree.IsCurrent &&
        WinApi.GetParent(SourceTab.Handle) == Window.Handle &&
        ExplorerNavigationAccess.ForegroundFrame() == Window.Handle;
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

    public static nint[] Tabs(nint window) => ExplorerWindowDiscovery.GetAllExplorerTabs(window).Take(MaxTabs + 1).ToArray();

    public static bool SameHandles(nint[] first, nint[] second) => first.Length == second.Length && new HashSet<nint>(first).SetEquals(second);

    /// <summary>Whether <paramref name="tree"/> is the navigation pane's tree control of <paramref name="window"/>.</summary>
    public static bool IsNavigationTree(nint tree, nint window) =>
        tree != 0 && window != 0 &&
        WinApi.IsWindowHasClassName(tree, "SysTreeView32") &&
        WinApi.IsWindowHasClassName(WinApi.GetParent(tree), "NamespaceTreeControl") &&
        WinApi.GetAncestor(tree, WinApi.GA_ROOT) == window;

    public static NavigationTabObservation Observe(nint window) => new(Tabs(window), ActiveTab(window));

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
        if (!isCurrent() || WinApi.GetParent(newTab) != window ||
            ActiveTab(window) != click.SourceTab.Handle || !SameHandles(Tabs(window), expected))
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
