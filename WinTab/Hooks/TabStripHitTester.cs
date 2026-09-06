using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Automation;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

using DrawingPoint = System.Drawing.Point;
using DrawingRectangle = System.Drawing.Rectangle;

internal sealed class TabStripHitTester : IDisposable
{
    private const int StaleBoundsMs = 1_500;
    private const int RectMatchSlop = 2;
    private readonly BackgroundRefreshCache<nint, TabStripBounds> _boundsCache = new(ComputeBounds);

    private sealed record TabStripBounds(DrawingRectangle[] TabRects, RECT WindowRect, long RefreshedAt, WindowIdentity Identity);

    public bool IsPointOnTabStrip(DrawingPoint screenPoint, nint explorerWindow)
    {
        if (!ExplorerWindowDiscovery.IsFileExplorerWindow(explorerWindow) ||
            !WinApi.GetWindowRect(explorerWindow, out var windowRect))
            return false;

        if (screenPoint.X < windowRect.Left || screenPoint.X >= windowRect.Right || screenPoint.Y < windowRect.Top)
            return false;

        if (_boundsCache.TryGet(explorerWindow, out var bounds) && bounds!.Identity.IsCurrent &&
            RectsApproxEqual(bounds.WindowRect, windowRect))
        {
            if (Environment.TickCount64 - bounds.RefreshedAt > StaleBoundsMs)
                ScheduleRefresh(explorerWindow);
            return bounds.TabRects.Any(rectangle => rectangle.Contains(screenPoint));
        }

        _boundsCache.Invalidate(explorerWindow);
        return false;
    }

    public void Refresh(nint explorerWindow) => _boundsCache.Invalidate(explorerWindow);
    public void ScheduleRefresh(nint explorerWindow) => _boundsCache.Request(explorerWindow);
    public void Forget(nint explorerWindow) => _boundsCache.Forget(explorerWindow);
    public void Clear() => _boundsCache.Clear();
    public void Dispose() => _boundsCache.Dispose();

    private static bool RectsApproxEqual(RECT first, RECT second) =>
        Math.Abs(first.Left - second.Left) <= RectMatchSlop &&
        Math.Abs(first.Top - second.Top) <= RectMatchSlop &&
        Math.Abs(first.Right - second.Right) <= RectMatchSlop &&
        Math.Abs(first.Bottom - second.Bottom) <= RectMatchSlop;

    private static TabStripBounds? ComputeBounds(nint explorerWindow)
    {
        if (!ExplorerWindowDiscovery.IsFileExplorerWindow(explorerWindow) ||
            !WinApi.GetWindowRect(explorerWindow, out var initialRect))
            return null;

        var identity = WindowIdentity.Capture(explorerWindow);
        if (!identity.IsCurrent)
            return null;
        var root = AutomationElement.FromHandle(explorerWindow);
        var tabItems = root.FindAll(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
        var tabRects = new List<DrawingRectangle>(tabItems.Count);
        foreach (AutomationElement tab in tabItems)
        {
            var rectangle = tab.Current.BoundingRectangle;
            if (!rectangle.IsEmpty && rectangle.Width >= 24 && rectangle.Height >= 12)
                tabRects.Add(new DrawingRectangle((int)rectangle.X, (int)rectangle.Y, (int)rectangle.Width, (int)rectangle.Height));
        }

        if (tabRects.Count == 0 || !identity.IsCurrent ||
            !WinApi.GetWindowRect(explorerWindow, out var finalRect) || !RectsApproxEqual(initialRect, finalRect))
            return null;

        return new TabStripBounds(tabRects.ToArray(), finalRect, Environment.TickCount64, identity);
    }
}
