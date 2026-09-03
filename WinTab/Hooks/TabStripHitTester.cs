using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

using DrawingPoint = System.Drawing.Point;
using DrawingRectangle = System.Drawing.Rectangle;

/// <summary>
/// Answers "is this screen point on an Explorer tab title?" using a cached UI Automation
/// snapshot per Explorer window, refreshed off the hook thread so mouse hooks never block.
/// </summary>
internal sealed class TabStripHitTester
{
    private const int AutomationRootTtlMs = 5_000;
    private const int StaleBoundsMs = 1_500;
    private const int RectMatchSlop = 2;

    private readonly ConcurrentDictionary<nint, TabStripBounds> _boundsCache = new();
    private readonly ConcurrentDictionary<nint, byte> _refreshInflight = new();
    private readonly ConcurrentDictionary<nint, (AutomationElement Element, long ExpiresAt)> _automationRootCache = new();

    private sealed record TabStripBounds(DrawingRectangle[] TabRects, RECT WindowRect, long RefreshedAt);

    public bool IsPointOnTabStrip(DrawingPoint screenPoint, nint explorerWindow)
    {
        if (!ExplorerWindowDiscovery.IsFileExplorerWindow(explorerWindow))
            return false;

        if (!WinApi.GetWindowRect(explorerWindow, out var winRect))
            return false;

        if (screenPoint.X < winRect.Left || screenPoint.X >= winRect.Right || screenPoint.Y < winRect.Top)
            return false;

        if (_boundsCache.TryGetValue(explorerWindow, out var bounds) && RectsApproxEqual(bounds.WindowRect, winRect))
        {
            if (Environment.TickCount64 - bounds.RefreshedAt > StaleBoundsMs)
                ScheduleRefresh(explorerWindow);

            return bounds.TabRects.Any(tabRect => tabRect.Contains(screenPoint.X, screenPoint.Y));
        }

        if (bounds != null)
            _boundsCache.TryRemove(explorerWindow, out _);

        ScheduleRefresh(explorerWindow);
        return false;
    }

    public void Refresh(nint explorerWindow)
    {
        InvalidateBounds(explorerWindow);
        if (ExplorerWindowDiscovery.IsFileExplorerWindow(explorerWindow))
            ScheduleRefresh(explorerWindow);
    }

    public void ScheduleRefresh(nint explorerWindow)
    {
        if (!_refreshInflight.TryAdd(explorerWindow, 0))
            return;

        Task.Run(() =>
        {
            try
            {
                ComputeBounds(explorerWindow);
            }
            catch
            {
                // 忽略边界刷新失败，后续扫描会再次尝试。
            }
            finally
            {
                _refreshInflight.TryRemove(explorerWindow, out _);
            }
        });
    }

    public void Forget(nint explorerWindow)
    {
        InvalidateBounds(explorerWindow);
        _automationRootCache.TryRemove(explorerWindow, out _);
    }

    public void Clear()
    {
        _boundsCache.Clear();
        _refreshInflight.Clear();
        _automationRootCache.Clear();
    }

    private void InvalidateBounds(nint explorerWindow)
    {
        _boundsCache.TryRemove(explorerWindow, out _);
        _refreshInflight.TryRemove(explorerWindow, out _);
    }

    private static bool RectsApproxEqual(RECT a, RECT b)
    {
        return Math.Abs(a.Left - b.Left) <= RectMatchSlop &&
               Math.Abs(a.Top - b.Top) <= RectMatchSlop &&
               Math.Abs(a.Right - b.Right) <= RectMatchSlop &&
               Math.Abs(a.Bottom - b.Bottom) <= RectMatchSlop;
    }

    private void ComputeBounds(nint explorerWindow)
    {
        if (!ExplorerWindowDiscovery.IsFileExplorerWindow(explorerWindow))
            return;

        var root = GetAutomationRoot(explorerWindow);
        if (root == null)
            return;

        var tabItems = root.FindAll(
            TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
        if (tabItems == null || tabItems.Count == 0)
            return;

        var tabRects = new List<DrawingRectangle>(tabItems.Count);
        foreach (AutomationElement tab in tabItems)
        {
            var rect = tab.Current.BoundingRectangle;
            if (rect.IsEmpty || rect.Width < 24 || rect.Height < 12)
                continue;

            tabRects.Add(new DrawingRectangle((int)rect.X, (int)rect.Y, (int)rect.Width, (int)rect.Height));
        }

        if (tabRects.Count == 0)
            return;

        if (!WinApi.GetWindowRect(explorerWindow, out var winSnapshot))
            return;

        _boundsCache[explorerWindow] = new TabStripBounds(tabRects.ToArray(), winSnapshot, Environment.TickCount64);
    }

    private AutomationElement? GetAutomationRoot(nint explorerWindow)
    {
        var now = Environment.TickCount64;
        if (_automationRootCache.TryGetValue(explorerWindow, out var entry) && entry.ExpiresAt > now)
            return entry.Element;

        try
        {
            var root = AutomationElement.FromHandle(explorerWindow);
            if (root != null)
                _automationRootCache[explorerWindow] = (root, now + AutomationRootTtlMs);

            return root;
        }
        catch
        {
            return null;
        }
    }
}
