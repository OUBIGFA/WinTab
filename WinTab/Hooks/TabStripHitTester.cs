using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

using DrawingPoint = System.Drawing.Point;
using DrawingRectangle = System.Drawing.Rectangle;

internal sealed class TabStripHitTester : IDisposable
{
    private const int StaleBoundsMs = 1_500;
    private const int RectMatchSlop = 2;
    private readonly BackgroundRefreshCache<nint, TabStripBounds> _boundsCache;

    public TabStripHitTester() : this(ExplorerTabAutomation.ReadTabs) { }

    internal TabStripHitTester(Func<nint, ExplorerTabAutomation.Tab[]> readTabs) =>
        _boundsCache = new(window => ComputeBounds(window, readTabs));

    /// <param name="TabRow">The band of the window that holds the tab titles, including the space beside them.</param>
    private sealed record TabStripBounds(DrawingRectangle[] TabRects, DrawingRectangle TabRow, RECT WindowRect, long RefreshedAt, WindowIdentity Identity);

    /// <summary>Whether the point is on one of the window's tab titles.</summary>
    public bool IsPointOnTabStrip(DrawingPoint screenPoint, nint explorerWindow)
    {
        if (!TryGetWindowRect(explorerWindow, out var windowRect))
            return false;

        if (screenPoint.X < windowRect.Left || screenPoint.X >= windowRect.Right || screenPoint.Y < windowRect.Top)
            return false;

        return TryGetBounds(explorerWindow, windowRect, out var bounds) &&
               bounds.TabRects.Any(rectangle => rectangle.Contains(screenPoint));
    }

    /// <summary>The band of the window that holds its tab titles, including the space beside them; false while it is not known yet.</summary>
    public bool TryGetTabRow(nint explorerWindow, out DrawingRectangle tabRow)
    {
        tabRow = DrawingRectangle.Empty;
        if (!TryGetWindowRect(explorerWindow, out var windowRect) || !TryGetBounds(explorerWindow, windowRect, out var bounds))
            return false;

        tabRow = bounds.TabRow;
        return true;
    }

    /// <summary>Resolves a cold hit test off the input hook without discarding the gesture that requested it.</summary>
    public async Task<DrawingRectangle?> GetTabRowAsync(nint explorerWindow, CancellationToken cancellationToken)
    {
        return await Helper.DoUntilConditionAsync<DrawingRectangle?>(() =>
            TryGetTabRow(explorerWindow, out var row) ? row : null,
            row => row.HasValue, 1_000, 5, cancellationToken);
    }

    private static bool TryGetWindowRect(nint explorerWindow, out RECT windowRect)
    {
        windowRect = default;
        return ExplorerWindowDiscovery.IsFileExplorerWindow(explorerWindow) &&
               WinApi.GetWindowRect(explorerWindow, out windowRect);
    }

    private bool TryGetBounds(nint explorerWindow, RECT windowRect, out TabStripBounds bounds)
    {
        if (_boundsCache.TryGet(explorerWindow, out var cached))
        {
            if (cached!.Identity.IsCurrent && RectsApproxEqual(cached.WindowRect, windowRect))
            {
                if (Environment.TickCount64 - cached.RefreshedAt > StaleBoundsMs)
                    ScheduleRefresh(explorerWindow);
                bounds = cached;
                return true;
            }
            _boundsCache.Invalidate(explorerWindow);
        }
        else
        {
            // Repeated cold lookups must not invalidate the computation already in flight.
            _boundsCache.Request(explorerWindow);
        }
        bounds = null!;
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

    private static TabStripBounds? ComputeBounds(nint explorerWindow, Func<nint, ExplorerTabAutomation.Tab[]> readTabs)
    {
        if (!ExplorerWindowDiscovery.IsFileExplorerWindow(explorerWindow) ||
            !WinApi.GetWindowRect(explorerWindow, out var initialRect))
            return null;

        var identity = WindowIdentity.Capture(explorerWindow);
        if (!identity.IsCurrent)
            return null;
        var tabs = readTabs(explorerWindow);
        var tabRects = new List<DrawingRectangle>(tabs.Length);
        foreach (var tab in tabs)
        {
            var rectangle = tab.Bounds;
            if (!rectangle.IsEmpty && rectangle.Width >= 24 && rectangle.Height >= 12)
                tabRects.Add(new DrawingRectangle((int)rectangle.X, (int)rectangle.Y, (int)rectangle.Width, (int)rectangle.Height));
        }

        if (tabRects.Count == 0 || !identity.IsCurrent ||
            !WinApi.GetWindowRect(explorerWindow, out var finalRect) || !RectsApproxEqual(initialRect, finalRect))
            return null;

        var tabRow = DrawingRectangle.FromLTRB(finalRect.Left, tabRects.Min(rectangle => rectangle.Top),
            finalRect.Right, tabRects.Max(rectangle => rectangle.Bottom));
        return new TabStripBounds(tabRects.ToArray(), tabRow, finalRect, Environment.TickCount64, identity);
    }
}
