using System;
using System.Drawing;
using System.Threading;

namespace WinTab.Hooks;

/// <summary>
/// Consumes non-overlapping down/up pairs using Windows' double-click settings. Recognition never waits
/// for a tab close or a bounds refresh; an unknown hit is validated by the worker without consuming input.
/// </summary>
internal sealed class ExplorerTabDoubleClickCloseController(IExplorerTabDoubleClickEnvironment environment)
{
    private ClickCandidate? _lastClickCandidate;
    private ClickCandidate? _pendingClose;
    private ClickCandidate? _recentClose;
    private bool _leftDown;
    private bool _firstClickReleased;
    private bool _suppressNextLeftUp;
    private int _interactionVersion;

    /// <summary>Unrelated input cancels queued actions, but the next pair at the same tab position does not.</summary>
    public int InteractionVersion => Volatile.Read(ref _interactionVersion);

    public MouseHookDecision HandleLeftMouseDown(Point currentPoint, long now)
    {
        // A repeated down without an up is not a second click. In particular, a drag must not arm a close.
        if (_leftDown)
            return MouseHookDecision.Native;
        _leftDown = true;

        var window = environment.IsEnabled ? environment.ResolveExplorerWindow(currentPoint) : 0;
        if (window == 0 || !environment.IsExplorerWindow(window))
        {
            CancelGesture();
            return MouseHookDecision.Native;
        }

        var hit = environment.IsPointOnTabStrip(currentPoint, window);
        if (hit == false)
        {
            CancelGesture();
            return MouseHookDecision.Native;
        }

        if (_recentClose == null || _recentClose.ExplorerWindow != window ||
            !IsWithinDoubleClickWindow(_recentClose, currentPoint, now))
        {
            Interlocked.Increment(ref _interactionVersion);
            _recentClose = null;
        }

        var previous = _lastClickCandidate;
        if (previous != null && _firstClickReleased && previous.ExplorerWindow == window &&
            IsWithinDoubleClickWindow(previous, currentPoint, now))
        {
            // Unknown means a cold/rebuilding strip, not a negative hit. Keep its native input intact,
            // then let the worker resolve the live tab before acting, for Explorer as well as Notepad.
            _suppressNextLeftUp = hit == true;
            _pendingClose = new ClickCandidate(window, currentPoint, now);
            _lastClickCandidate = null;
            _firstClickReleased = false;
            return new MouseHookDecision(_suppressNextLeftUp, null);
        }

        _lastClickCandidate = new ClickCandidate(window, currentPoint, now);
        _firstClickReleased = false;
        return MouseHookDecision.Native;
    }

    public MouseHookDecision HandleLeftMouseUp(long now, Point? currentPoint = null)
    {
        if (currentPoint is { } point)
            HandleMouseMove(point);
        _leftDown = false;
        _firstClickReleased = _lastClickCandidate != null;

        var handled = _suppressNextLeftUp;
        _suppressNextLeftUp = false;
        var pending = _pendingClose;
        _pendingClose = null;
        if (pending == null || !environment.IsEnabled)
            return new MouseHookDecision(handled, null);

        _recentClose = pending with { Tick = now };
        return new MouseHookDecision(handled, new ExplorerTabCloseRequest(pending.ExplorerWindow, pending.Point));
    }

    public void HandleMouseMove(Point point)
    {
        if (!_leftDown)
            return;
        var pressed = _pendingClose ?? _lastClickCandidate;
        if (pressed != null && !IsWithinDoubleClickDistance(pressed.Point, point))
            CancelGesture();
    }

    /// <summary>Cancel an interrupted pair, but still consume the up matching an already consumed down.</summary>
    public void CancelGesture()
    {
        Interlocked.Increment(ref _interactionVersion);
        _lastClickCandidate = null;
        _pendingClose = null;
        _recentClose = null;
        _firstClickReleased = false;
    }

    public void Reset()
    {
        CancelGesture();
        _leftDown = false;
        _suppressNextLeftUp = false;
    }

    private bool IsWithinDoubleClickWindow(ClickCandidate previous, Point currentPoint, long now) =>
        now >= previous.Tick && now - previous.Tick <= environment.DoubleClickTimeMs &&
        IsWithinDoubleClickDistance(previous.Point, currentPoint);

    private bool IsWithinDoubleClickDistance(Point previousPoint, Point currentPoint) =>
        IsWithinDoubleClickDistance(previousPoint, currentPoint, environment.DoubleClickWidth, environment.DoubleClickHeight);

    internal static bool IsWithinDoubleClickDistance(Point previousPoint, Point currentPoint, int width, int height)
    {
        // SM_CXDOUBLECLK/SM_CYDOUBLECLK are the full rectangle sizes, not a radius. Use wide arithmetic
        // so coordinates on the virtual desktop cannot overflow when calculating the rectangle edges.
        width = Math.Max(1, width);
        height = Math.Max(1, height);
        var left = (long)previousPoint.X - width / 2;
        var top = (long)previousPoint.Y - height / 2;
        return currentPoint.X >= left && currentPoint.X < left + width &&
               currentPoint.Y >= top && currentPoint.Y < top + height;
    }

    private sealed record ClickCandidate(nint ExplorerWindow, Point Point, long Tick);
}

/// <summary>The environment resolves windows in the feature's scope: Explorer frames, plus Notepad when included.</summary>
internal interface IExplorerTabDoubleClickEnvironment
{
    bool IsEnabled { get; }
    int DoubleClickTimeMs { get; }
    int DoubleClickWidth { get; }
    int DoubleClickHeight { get; }
    nint ResolveExplorerWindow(Point point);
    bool IsExplorerWindow(nint explorerWindow);
    /// <summary>Null means unavailable bounds; false is a confirmed non-tab hit.</summary>
    bool? IsPointOnTabStrip(Point point, nint explorerWindow);
}

internal readonly record struct ExplorerTabCloseRequest(nint ExplorerWindow, Point Point);

internal readonly record struct MouseHookDecision(bool Handled, ExplorerTabCloseRequest? CloseRequest)
{
    public static MouseHookDecision Native { get; } = new(false, null);
    public static MouseHookDecision HandledOnly { get; } = new(true, null);
}
