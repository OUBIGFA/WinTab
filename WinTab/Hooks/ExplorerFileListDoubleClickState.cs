using System;
using System.Drawing;
using WinTab.Helpers;

namespace WinTab.Hooks;

internal sealed record ExplorerFileListClick(
    WindowIdentity Window,
    WindowIdentity Tab,
    WindowIdentity View,
    Point Point,
    string Location,
    long Tick)
{
    internal WinAPI.RECT ViewBounds { get; init; }

    internal bool IsCurrent() =>
        Window.IsCurrent && Tab.IsCurrent && View.IsCurrent &&
        WinAPI.WinApi.GetParent(Tab.Handle) == Window.Handle &&
        ExplorerNavigationAccess.ActiveTab(Window.Handle) == Tab.Handle;
}

internal sealed record ExplorerFileListDoubleClickRequest(ExplorerFileListClick First, ExplorerFileListClick Second);

/// <summary>
/// Recognizes non-overlapping left-button pairs without consuming native Explorer input. Both click
/// snapshots travel with the request so background hit testing and navigation can reject stale gestures.
/// </summary>
internal sealed class ExplorerFileListDoubleClickState(
    Func<int> doubleClickTimeMs,
    Func<int> doubleClickWidth,
    Func<int> doubleClickHeight)
{
    private ExplorerFileListClick? _first;
    private ExplorerFileListClick? _second;
    private bool _leftDown;
    private bool _firstReleased;

    internal bool HasPair => _second != null;

    public ExplorerFileListClick? HandleLeftDown(ExplorerFileListClick? click)
    {
        if (_leftDown) { Reset(); return null; }
        var firstReleased = _firstReleased;
        _leftDown = true;
        _second = null;
        _firstReleased = false;

        if (click == null)
        {
            _first = null;
            return null;
        }

        var previous = _first;
        if (previous != null && firstReleased && SameTarget(previous, click) &&
            WithinTime(previous, click) && WithinDistance(previous.Point, click.Point))
        {
            _first = previous;
            _second = click;
            return null;
        }

        _first = click;
        return click;
    }

    public ExplorerFileListDoubleClickRequest? HandleLeftUp(Point point, long now)
    {
        if (!_leftDown) return null;
        _leftDown = false;

        if (_second is { } second)
        {
            var completedFirst = _first!;
            _first = null;
            _second = null;
            _firstReleased = false;
            return WithinDistance(second.Point, point) && now >= second.Tick && now - second.Tick <= Math.Max(1, doubleClickTimeMs())
                ? new ExplorerFileListDoubleClickRequest(completedFirst, second)
                : null;
        }

        if (_first is { } first && WithinDistance(first.Point, point) && now >= first.Tick &&
            now - first.Tick <= Math.Max(1, doubleClickTimeMs()))
        {
            _firstReleased = true;
        }
        else
        {
            Reset();
        }
        return null;
    }

    public void Reset()
    {
        _first = null;
        _second = null;
        _leftDown = false;
        _firstReleased = false;
    }

    private static bool SameTarget(ExplorerFileListClick first, ExplorerFileListClick second) =>
        first.Window == second.Window && first.Tab == second.Tab && first.View == second.View &&
        first.ViewBounds.Equals(second.ViewBounds) && StringComparer.OrdinalIgnoreCase.Equals(first.Location, second.Location);

    private bool WithinTime(ExplorerFileListClick first, ExplorerFileListClick second) =>
        second.Tick >= first.Tick && second.Tick - first.Tick <= Math.Max(1, doubleClickTimeMs());

    internal bool WithinDistance(Point first, Point second) =>
        ExplorerTabDoubleClickCloseController.IsWithinDoubleClickDistance(first, second,
            Math.Max(1, doubleClickWidth()), Math.Max(1, doubleClickHeight()));
}
