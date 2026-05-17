using System;
using System.Drawing;

namespace WinTab.Hooks;

internal sealed class ExplorerTabDoubleClickCloseController(IExplorerTabDoubleClickEnvironment environment)
{
    private const int CloseChainFallbackMs = 1_500;

    private ClickCandidate? _lastClickCandidate;
    private ClickCandidate? _pendingNativeClose;
    private ClickCandidate? _recentNativeClose;
    private bool _suppressNextLeftUp;

    public MouseHookDecision HandleLeftMouseDown(Point currentPoint, long now)
    {
        if (!environment.IsEnabled)
        {
            ResetClickState();
            return MouseHookDecision.Native;
        }

        var previous = _lastClickCandidate;
        var explorerWindow = environment.ResolveExplorerWindow(currentPoint);
        if (explorerWindow == 0)
        {
            if (previous != null &&
                IsWithinDoubleClickWindow(previous, currentPoint, now) &&
                environment.IsExplorerWindow(previous.ExplorerWindow))
            {
                explorerWindow = previous.ExplorerWindow;
            }
            else if (TryGetCloseChainWindow(currentPoint, now, out var chainedWindow))
            {
                explorerWindow = chainedWindow;
            }
            else
            {
                ResetClickState();
                return MouseHookDecision.Native;
            }
        }

        if (!environment.IsExplorerWindow(explorerWindow))
        {
            ResetClickState();
            return MouseHookDecision.Native;
        }

        var onTabStrip = IsPointOnTabStrip(currentPoint, explorerWindow, now);
        if (previous != null &&
            previous.ExplorerWindow == explorerWindow &&
            IsWithinDoubleClickWindow(previous, currentPoint, now) &&
            onTabStrip &&
            previous.OnTabStrip)
        {
            _suppressNextLeftUp = true;
            _pendingNativeClose = new ClickCandidate(explorerWindow, currentPoint, now) { OnTabStrip = true };
            _lastClickCandidate = null;
            return MouseHookDecision.HandledOnly;
        }

        _lastClickCandidate = onTabStrip
            ? new ClickCandidate(explorerWindow, currentPoint, now) { OnTabStrip = true }
            : null;

        return MouseHookDecision.Native;
    }

    public MouseHookDecision HandleLeftMouseUp(long now)
    {
        if (!_suppressNextLeftUp)
            return MouseHookDecision.Native;

        _suppressNextLeftUp = false;

        var pending = _pendingNativeClose;
        _pendingNativeClose = null;
        if (pending == null)
            return MouseHookDecision.HandledOnly;

        _recentNativeClose = new ClickCandidate(pending.ExplorerWindow, pending.Point, now) { OnTabStrip = true };
        return new MouseHookDecision(
            true,
            new ExplorerTabCloseRequest(pending.ExplorerWindow, pending.Point));
    }

    private bool IsPointOnTabStrip(Point currentPoint, nint explorerWindow, long now)
    {
        return environment.IsPointOnTabStrip(currentPoint, explorerWindow) ||
               IsWithinCloseChain(explorerWindow, currentPoint, now);
    }

    private bool TryGetCloseChainWindow(Point currentPoint, long now, out nint explorerWindow)
    {
        explorerWindow = 0;
        if (_recentNativeClose == null || !IsWithinCloseChain(_recentNativeClose.ExplorerWindow, currentPoint, now))
            return false;

        explorerWindow = _recentNativeClose.ExplorerWindow;
        return true;
    }

    private bool IsWithinCloseChain(nint explorerWindow, Point currentPoint, long now)
    {
        var anchor = _recentNativeClose;
        if (anchor == null ||
            anchor.ExplorerWindow != explorerWindow ||
            now - anchor.Tick > CloseChainFallbackMs)
        {
            return false;
        }

        if (!environment.IsExplorerWindow(anchor.ExplorerWindow))
        {
            _recentNativeClose = null;
            return false;
        }

        return IsWithinDoubleClickDistance(anchor.Point, currentPoint);
    }

    private bool IsWithinDoubleClickWindow(ClickCandidate previous, Point currentPoint, long now)
    {
        return now - previous.Tick <= Math.Max(200, environment.DoubleClickTimeMs) &&
               IsWithinDoubleClickDistance(previous.Point, currentPoint);
    }

    private bool IsWithinDoubleClickDistance(Point previousPoint, Point currentPoint)
    {
        var maxX = Math.Max(4, environment.DoubleClickWidth);
        var maxY = Math.Max(4, environment.DoubleClickHeight);

        return Math.Abs(currentPoint.X - previousPoint.X) <= maxX &&
               Math.Abs(currentPoint.Y - previousPoint.Y) <= maxY;
    }

    private void ResetClickState()
    {
        _lastClickCandidate = null;
        if (!_suppressNextLeftUp)
            _pendingNativeClose = null;
    }

    private sealed record ClickCandidate(nint ExplorerWindow, Point Point, long Tick)
    {
        public bool OnTabStrip { get; init; }
    }
}

internal interface IExplorerTabDoubleClickEnvironment
{
    bool IsEnabled { get; }
    int DoubleClickTimeMs { get; }
    int DoubleClickWidth { get; }
    int DoubleClickHeight { get; }
    nint ResolveExplorerWindow(Point point);
    bool IsExplorerWindow(nint explorerWindow);
    bool IsPointOnTabStrip(Point point, nint explorerWindow);
}

internal readonly record struct ExplorerTabCloseRequest(nint ExplorerWindow, Point Point);

internal readonly record struct MouseHookDecision(bool Handled, ExplorerTabCloseRequest? CloseRequest)
{
    public static MouseHookDecision Native { get; } = new(false, null);
    public static MouseHookDecision HandledOnly { get; } = new(true, null);
}
