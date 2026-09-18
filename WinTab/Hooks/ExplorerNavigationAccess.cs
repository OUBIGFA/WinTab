using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using WinTab.Helpers;
using WinTab.Interop;
using WinTab.WinAPI;

namespace WinTab.Hooks;

internal sealed record NavigationItem(int Id, string Name, Rectangle LabelBounds, Rectangle HitBounds);

internal sealed record ExplorerNavigationSnapshot(
    WindowIdentity Window, WindowIdentity SourceTab, WindowIdentity Tree,
    RECT WindowRect, RECT TreeRect, WindowIdentity[] TabIdentities,
    NavigationTabSet Tabs, NavigationItem[] Items, long CreatedAt)
{
    public bool Matches(nint window, nint tree, nint[] tabs) =>
        Window.Handle == window && Tree.Handle == tree && Window.IsCurrent && SourceTab.IsCurrent && Tree.IsCurrent &&
        ExplorerNavigationAccess.ActiveTab(window) == SourceTab.Handle &&
        TabIdentities.All(identity => identity.IsCurrent) && ExplorerNavigationAccess.SameHandles(Tabs.Handles, tabs) &&
        HasSameBounds();

    public bool HasSameBounds() =>
        WinApi.GetWindowRect(Window.Handle, out var frame) && frame.Equals(WindowRect) &&
        WinApi.GetWindowRect(Tree.Handle, out var tree) && tree.Equals(TreeRect);
}

/// <summary>All accessibility work is performed off the low-level hook thread.</summary>
internal static class ExplorerNavigationAccess
{
    internal const int MaxTabs = 128;
    private const uint ObjIdClient = 0xFFFFFFFC;
    private const int OutlineItemRole = 0x24;
    private const int UnavailableOrInvisible = 0x1 | 0x8000 | 0x10000;

    public static nint ActiveTab(nint window) => WinApi.FindWindowEx(window, 0, "ShellTabWindowClass", null);

    public static nint NavigationTree(nint tab)
    {
        if (tab == 0) return 0;
        nint result = 0;
        var examined = 0;
        EnumChildWindows(tab, (child, _) =>
        {
            if (++examined > 256) return false;
            if (WinApi.IsWindowHasClassName(child, "SysTreeView32") &&
                WinApi.IsWindowHasClassName(WinApi.GetParent(child), "NamespaceTreeControl"))
            {
                result = child;
                return false;
            }
            return true;
        }, 0);
        return result;
    }

    public static nint[] Tabs(nint window) => ExplorerWindowDiscovery.GetAllExplorerTabs(window).Take(MaxTabs + 1).ToArray();
    public static bool SameHandles(nint[] first, nint[] second) => first.Length == second.Length && new HashSet<nint>(first).SetEquals(second);

    public static ExplorerNavigationSnapshot? Capture(nint window)
    {
        if (!ExplorerWindowDiscovery.IsShownExplorerWindow(window)) return null;
        var identity = WindowIdentity.Capture(window);
        var tab = ActiveTab(window);
        var tree = NavigationTree(tab);
        var handles = Tabs(window);
        if (tree == 0 || handles.Length is 0 or > MaxTabs ||
            !WinApi.GetWindowRect(window, out var windowRect) || !WinApi.GetWindowRect(tree, out var treeRect))
            return null;
        var sourceIdentity = WindowIdentity.Capture(tab);
        var treeIdentity = WindowIdentity.Capture(tree);
        var tabIdentities = handles.Select(WindowIdentity.Capture).ToArray();
        var automationIds = TabElements(window).Select(Id).ToArray();
        if (automationIds.Length != handles.Length || automationIds.Distinct().Count() != automationIds.Length)
            return null;
        var items = ReadItems(tree);
        var snapshot = new ExplorerNavigationSnapshot(identity, sourceIdentity, treeIdentity, windowRect, treeRect,
            tabIdentities, new NavigationTabSet(handles, automationIds), items, Environment.TickCount64);
        return snapshot.Matches(window, tree, Tabs(window)) ? snapshot : null;
    }

    public static AutomationElement[] TabElements(nint window)
    {
        // Modern Explorer keeps the tab strip in its own HWND. Searching the whole frame also walks
        // every inactive folder view (potentially thousands of files) and adds visible activation lag.
        var tabCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
        var bridges = WinApi.FindAllWindowsEx("Microsoft.UI.Content.DesktopChildSiteBridge", window)
            .Concat(WinApi.FindAllWindowsEx("Windows.UI.Composition.DesktopWindowContentBridge", window));
        foreach (var bridge in bridges)
        {
            var root = AutomationElement.FromHandle(bridge);
            var tabs = root.FindAll(TreeScope.Descendants, tabCondition);
            if (tabs.Count > 0)
                return tabs.Cast<AutomationElement>().Take(MaxTabs + 1).ToArray();
        }
        var fallbackRoot = AutomationElement.FromHandle(window);
        var fallbackTabs = fallbackRoot.FindAll(TreeScope.Descendants, tabCondition);
        return fallbackTabs.Cast<AutomationElement>().Take(MaxTabs + 1).ToArray();
    }

    public static string Id(AutomationElement element) => string.Join(":", element.GetRuntimeId());

    public static NavigationTabObservation Observe(nint window, NavigationTabSet before)
    {
        var handles = Tabs(window);
        // Native creation may precede UIA publication. Avoid accessibility scans while no tab has appeared.
        var ids = SameHandles(handles, before.Handles) ? before.AutomationIds : TabElements(window).Select(Id).ToArray();
        return new NavigationTabObservation(new NavigationTabSet(handles, ids), ActiveTab(window));
    }

    public static bool Select(ExplorerNavigationSnapshot snapshot, NavigationItem item, Point point,
        string targetId, Func<bool> isCurrent)
    {
        if (!isCurrent() || !HitStillMatches(snapshot, item, point)) return false;
        var handles = Tabs(snapshot.Window.Handle);
        if (!NavigationTabActivation.TryGetOnlyAddition(snapshot.Tabs.Handles, handles, out var targetHandle)) return false;
        var targetIdentity = WindowIdentity.Capture(targetHandle);
        if (!targetIdentity.IsCurrent || WinApi.GetParent(targetHandle) != snapshot.Window.Handle) return false;
        var elements = TabElements(snapshot.Window.Handle);
        if (!NavigationTabActivation.TryGetOnlyAddition(snapshot.Tabs.AutomationIds, elements.Select(Id).ToArray(), out var id) ||
            id != targetId) return false;
        var target = elements.Single(element => Id(element) == targetId);
        if (!target.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var pattern)) return false;
        // UIA queries can block. Recheck after them, immediately before the only mutating operation.
        if (!isCurrent() || !targetIdentity.IsCurrent || WinApi.GetParent(targetHandle) != snapshot.Window.Handle ||
            ActiveTab(snapshot.Window.Handle) != snapshot.SourceTab.Handle ||
            !SameHandles(handles, Tabs(snapshot.Window.Handle))) return false;
        ((SelectionItemPattern)pattern).Select();
        return true;
    }

    public static bool HitStillMatches(ExplorerNavigationSnapshot snapshot, NavigationItem expected, Point point)
    {
        if (!snapshot.Tree.IsCurrent || !snapshot.HasSameBounds()) return false;
        var accessible = AccessibleTree(snapshot.Tree.Handle);
        try
        {
            var child = accessible.accHitTest(point.X, point.Y);
            try
            {
                if (child is not int id || id != expected.Id) return false;
                var current = ReadItem(accessible, id, snapshot.Tree.Handle);
                return current != null && current.Name == expected.Name && current.LabelBounds == expected.LabelBounds &&
                    current.HitBounds.Contains(point);
            }
            finally
            {
                if (child != null && Marshal.IsComObject(child))
                    Marshal.ReleaseComObject(child);
            }
        }
        finally { Marshal.ReleaseComObject(accessible); }
    }

    private static NavigationItem[] ReadItems(nint tree)
    {
        var accessible = AccessibleTree(tree);
        try
        {
            var count = accessible.get_accChildCount();
            if (count is <= 0 or > 512) return [];
            var children = new object[count];
            Marshal.ThrowExceptionForHR(AccessibleChildren(accessible, 0, count, children, out var obtained));
            var items = new List<NavigationItem>();
            for (var i = 0; i < obtained; i++)
            {
                // SysTreeView32 exposes stable integer child IDs. Do not guess identities for other providers.
                if (children[i] is int id)
                {
                    var item = ReadItem(accessible, id, tree);
                    if (item != null) items.Add(item);
                }
                else if (children[i] != null && Marshal.IsComObject(children[i]))
                    Marshal.ReleaseComObject(children[i]);
            }
            return items.ToArray();
        }
        finally { Marshal.ReleaseComObject(accessible); }
    }

    private static NavigationItem? ReadItem(IAccessible accessible, int id, nint tree)
    {
        if (id <= 0 || Convert.ToInt32(accessible.get_accRole(id)) != OutlineItemRole ||
            (Convert.ToInt32(accessible.get_accState(id)) & UnavailableOrInvisible) != 0) return null;
        accessible.accLocation(out var left, out var top, out var width, out var height, id);
        if (width <= 0 || height <= 0 || !WinApi.GetWindowRect(tree, out var bounds)) return null;
        var label = new Rectangle(left, top, width, height);
        var dpi = Math.Max(96u, WinApi.GetDpiForWindow(tree));
        // MSAA gives the label rectangle. Include its adjacent small icon, not the expansion glyph to its left.
        var iconWidth = GetSystemMetricsForDpi(WinApi.SM_CXSMICON, dpi) + (int)(4 * dpi / 96);
        var hit = Rectangle.Intersect(new Rectangle(left - iconWidth, top, width + iconWidth, height),
            new Rectangle(bounds.Left, bounds.Top, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top));
        return hit.IsEmpty ? null : new NavigationItem(id, accessible.get_accName(id), label, hit);
    }

    private static IAccessible AccessibleTree(nint tree)
    {
        var guid = typeof(IAccessible).GUID;
        Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(tree, ObjIdClient, ref guid, out var accessible));
        return accessible;
    }

    private delegate bool EnumChildCallback(nint window, nint data);
    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(nint parent, EnumChildCallback callback, nint data);
    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(nint window, uint objectId, ref Guid interfaceId,
        [MarshalAs(UnmanagedType.Interface)] out IAccessible accessible);
    [DllImport("oleacc.dll")]
    private static extern int AccessibleChildren(IAccessible accessible, int start, int count,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] object[] children, out int obtained);
    [DllImport("user32.dll")]
    private static extern int GetSystemMetricsForDpi(int index, uint dpi);
}
