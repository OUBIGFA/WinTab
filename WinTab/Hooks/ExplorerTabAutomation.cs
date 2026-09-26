using System;
using System.Windows.Automation;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Reads only Explorer's tab control, never the file views below ShellTabWindowClass. A single cached
/// UIA request returns the rectangles and selection in visual order without a round trip per property.
/// No AutomationElement or selection is retained across tab moves, window recreation or threads.
/// </summary>
internal static class ExplorerTabAutomation
{
    private static readonly Condition TabCondition =
        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab);
    private static readonly Condition TabItemCondition =
        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
    private static readonly Condition CloseButtonCondition =
        new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
            new PropertyCondition(AutomationElement.AutomationIdProperty, "CloseButton"));

    internal readonly record struct Tab(System.Windows.Rect Bounds, bool Selected, string Title = "", string Id = "");

    public static Tab[] ReadTabs(nint window) => ReadTabs(window, includeIdentity: false);

    /// <summary>Session snapshots additionally need titles and stable UIA identities, never live UIA objects.</summary>
    public static Tab[] ReadSessionTabs(nint window) => ReadTabs(window, includeIdentity: true);

    private static Tab[] ReadTabs(nint window, bool includeIdentity)
    {
        var tabControl = FindTabControl(window);
        if (tabControl == null)
            return [];

        var cache = new CacheRequest { TreeScope = TreeScope.Element, AutomationElementMode = AutomationElementMode.None };
        cache.Add(AutomationElement.BoundingRectangleProperty);
        cache.Add(SelectionItemPattern.IsSelectedProperty);
        if (includeIdentity)
        {
            cache.Add(AutomationElement.NameProperty);
            cache.Add(AutomationElement.RuntimeIdProperty);
        }
        using (cache.Activate())
        {
            var elements = tabControl.FindAll(TreeScope.Descendants, TabItemCondition);
            var tabs = new Tab[elements.Count];
            for (var index = 0; index < elements.Count; index++)
            {
                var element = elements[index];
                var id = includeIdentity && element.GetCachedPropertyValue(AutomationElement.RuntimeIdProperty) is int[] runtimeId
                    ? string.Join(".", runtimeId) : string.Empty;
                tabs[index] = new Tab(element.Cached.BoundingRectangle,
                    element.GetCachedPropertyValue(SelectionItemPattern.IsSelectedProperty) is true,
                    includeIdentity ? element.Cached.Name : string.Empty, id);
            }
            return tabs;
        }
    }

    private static AutomationElement? FindTabControl(nint window)
    {
        // Current WinUI and earlier Windows 11 XAML islands use different native host classes.
        var host = WinApi.FindWindowEx(window, 0, "Microsoft.UI.Content.DesktopChildSiteBridge", null);
        if (host == 0)
            host = WinApi.FindWindowEx(window, 0, "Windows.UI.Composition.DesktopWindowContentBridge", null);
        if (host != 0)
            return AutomationElement.FromHandle(host).FindFirst(TreeScope.Descendants, TabCondition);

        // Other Explorer layouts can expose the tab control through the frame instead of an island.
        // Search those frame children, but never enter a tab's folder/navigation/preview subtree.
        var children = AutomationElement.FromHandle(window).FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement child in children)
        {
            if (child.Current.ClassName == "ShellTabWindowClass") continue;
            if (child.Current.ControlType == ControlType.Tab) return child;
            var tabControl = child.FindFirst(TreeScope.Descendants, TabCondition);
            if (tabControl != null) return tabControl;
        }
        return null;
    }

    public static bool TryCloseSelectedTab(nint window, string expectedId, Func<bool> stillInitialTab)
    {
        if (string.IsNullOrEmpty(expectedId))
            return false;
        var tabControl = FindTabControl(window);
        if (tabControl == null)
            return false;
        var items = tabControl.FindAll(TreeScope.Descendants, TabItemCondition);
        AutomationElement? selected = null;
        foreach (AutomationElement item in items)
        {
            if (item.GetCurrentPropertyValue(SelectionItemPattern.IsSelectedProperty) is not true)
                continue;
            if (selected != null)
                return false;
            selected = item;
        }
        if (selected == null || !string.Equals(string.Join(".", selected.GetRuntimeId()), expectedId, StringComparison.Ordinal))
            return false;
        var button = selected.FindFirst(TreeScope.Descendants, CloseButtonCondition);
        if (button == null || !button.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern))
            return false;
        // Reading a remote pattern may block; recheck the click/restore after that call, immediately before
        // invoking the exact tab's button. Never let a delayed UIA read act on a retired restoration.
        if (!stillInitialTab())
            return false;
        ((InvokePattern)pattern).Invoke();
        return true;
    }

    public static bool TryReadSelection(nint window, out int selectedIndex, out int count)
    {
        var tabs = ReadTabs(window);
        count = tabs.Length;
        selectedIndex = Array.FindIndex(tabs, tab => tab.Selected);
        return selectedIndex >= 0;
    }
}
