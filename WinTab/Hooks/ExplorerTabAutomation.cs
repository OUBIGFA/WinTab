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

    internal readonly record struct Tab(System.Windows.Rect Bounds, bool Selected);

    public static Tab[] ReadTabs(nint window)
    {
        var tabControl = FindTabControl(window);
        if (tabControl == null)
            return [];

        var cache = new CacheRequest { TreeScope = TreeScope.Element, AutomationElementMode = AutomationElementMode.None };
        cache.Add(AutomationElement.BoundingRectangleProperty);
        cache.Add(SelectionItemPattern.IsSelectedProperty);
        using (cache.Activate())
        {
            var elements = tabControl.FindAll(TreeScope.Descendants, TabItemCondition);
            var tabs = new Tab[elements.Count];
            for (var index = 0; index < elements.Count; index++)
                tabs[index] = new Tab(elements[index].Cached.BoundingRectangle,
                    elements[index].GetCachedPropertyValue(SelectionItemPattern.IsSelectedProperty) is true);
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

    public static bool TryReadSelection(nint window, out int selectedIndex, out int count)
    {
        var tabs = ReadTabs(window);
        count = tabs.Length;
        selectedIndex = Array.FindIndex(tabs, tab => tab.Selected);
        return selectedIndex >= 0;
    }
}
