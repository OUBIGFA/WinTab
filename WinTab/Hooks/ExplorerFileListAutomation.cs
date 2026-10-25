using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Automation;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Captures Explorer's active DirectUI file view with native HWNDs, then uses UIA off the input thread to
/// positively identify its empty background. Unknown providers and controls fail closed.
/// </summary>
internal static class ExplorerFileListAutomation
{
    internal static bool TryCaptureView(Point point, out WindowIdentity window, out WindowIdentity tab,
        out WindowIdentity view)
    {
        window = default;
        tab = default;
        view = default;

        var hit = WinApi.WindowFromPoint(point);
        if (hit == 0 || !TryFindFileView(hit, out var viewHandle, out var tabHandle))
            return false;

        var windowHandle = WinApi.GetAncestor(tabHandle, WinApi.GA_ROOT);
        if (windowHandle == 0 || WinApi.GetParent(tabHandle) != windowHandle ||
            !ExplorerWindowDiscovery.IsShownExplorerWindow(windowHandle) ||
            ExplorerNavigationAccess.ActiveTab(windowHandle) != tabHandle)
            return false;

        var capturedWindow = WindowIdentity.Capture(windowHandle);
        var capturedTab = WindowIdentity.Capture(tabHandle);
        var capturedView = WindowIdentity.Capture(viewHandle);
        if (!capturedWindow.IsCurrent || !capturedTab.IsCurrent || !capturedView.IsCurrent)
            return false;

        window = capturedWindow;
        tab = capturedTab;
        view = capturedView;
        return true;
    }

    internal static bool IsCurrent(ExplorerFileListClick click)
    {
        if (!click.IsCurrent() || !WinApi.GetWindowRect(click.View.Handle, out var bounds) ||
            !bounds.Equals(click.ViewBounds))
            return false;
        if (!TryFindFileView(click.View.Handle, out var view, out var tab) ||
            view != click.View.Handle || tab != click.Tab.Handle)
            return false;

        var hit = WinApi.WindowFromPoint(click.Point);
        return hit != 0 && (hit == click.View.Handle || IsDescendant(hit, click.View.Handle));
    }

    /// <summary>Runs on a worker thread. Only UIItemsView/UIGroupItem background is accepted.</summary>
    internal static bool IsEmptyArea(ExplorerFileListClick click)
    {
        if (!IsCurrent(click))
            return false;

        try
        {
            var element = AutomationElement.FromPoint(new System.Windows.Point(click.Point.X, click.Point.Y));
            if (element == null)
                return false;

            return IsBackgroundPath(ReadAncestry(element), click.View.Handle) && IsCurrent(click);
        }
        catch (Exception exception) when (exception is ElementNotAvailableException or InvalidOperationException or System.Runtime.InteropServices.COMException)
        {
            ExplorerDebugLog.Write($"File-list hit test unavailable: {exception.GetType().Name}");
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"File-list hit test failed: {exception.GetType().Name}:{exception.Message}");
        }
        return false;
    }

    // Read lazily: stop walking as soon as the native view boundary or an occupied control is found.
    private static IEnumerable<(ControlType Type, string ClassName, nint Window)> ReadAncestry(AutomationElement element)
    {
        var walker = TreeWalker.RawViewWalker;
        for (var current = element; current != null; current = walker.GetParent(current))
        {
            var properties = current.Current;
            yield return (properties.ControlType, properties.ClassName, properties.NativeWindowHandle);
        }
    }

    /// <summary>Shared hit policy, exercised with element snapshots without a desktop UIA query.</summary>
    internal static bool IsBackgroundPath(IEnumerable<(ControlType Type, string ClassName, nint Window)> path, nint view)
    {
        if (view == 0) return false;
        var depth = 0;
        var foundBackground = false;
        foreach (var element in path)
        {
            if (++depth > 64 || IsOccupiedOrInteractive(element.Type)) return false;
            if (element.ClassName is "UIItemsView" or "UIGroupItem") foundBackground = true;
            if (element.Window == view) return foundBackground;
        }
        return false;
    }

    private static bool IsOccupiedOrInteractive(ControlType type) =>
        type == ControlType.ListItem || type == ControlType.DataItem || type == ControlType.TreeItem ||
        type == ControlType.Button || type == ControlType.Hyperlink || type == ControlType.Edit ||
        type == ControlType.ComboBox || type == ControlType.CheckBox || type == ControlType.ScrollBar ||
        type == ControlType.Thumb || type == ControlType.Header || type == ControlType.HeaderItem ||
        type == ControlType.Menu || type == ControlType.MenuItem || type == ControlType.Text;

    private static bool TryFindFileView(nint hit, out nint view, out nint tab)
    {
        view = 0;
        tab = 0;
        var current = hit;
        for (var depth = 0; current != 0 && depth < 64; depth++)
        {
            if (WinApi.IsWindowHasClassName(current, "DirectUIHWND") &&
                WinApi.IsWindowHasClassName(WinApi.GetParent(current), "SHELLDLL_DefView"))
            {
                view = current;
                break;
            }
            current = WinApi.GetParent(current);
        }
        if (view == 0)
            return false;

        current = WinApi.GetParent(view);
        for (var depth = 0; current != 0 && depth < 64; depth++)
        {
            if (WinApi.IsWindowHasClassName(current, "ShellTabWindowClass"))
            {
                tab = current;
                return true;
            }
            current = WinApi.GetParent(current);
        }
        return false;
    }

    private static bool IsDescendant(nint handle, nint ancestor)
    {
        var current = handle;
        for (var depth = 0; current != 0 && depth < 64; depth++)
        {
            if (current == ancestor)
                return true;
            current = WinApi.GetParent(current);
        }
        return false;
    }
}
