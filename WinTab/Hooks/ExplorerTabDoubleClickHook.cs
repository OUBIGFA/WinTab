using System;
using System.Drawing;
using System.Threading;
using H.Hooks;
using WinTab.Helpers;
using WinTab.Managers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

public sealed class ExplorerTabDoubleClickHook : IHook
{
    private const int SM_CXDOUBLECLK = 36;
    private const int SM_CYDOUBLECLK = 37;
    private const uint GA_ROOT = 2;

    private readonly ExplorerWatcher _explorerWatcher;
    private readonly LowLevelMouseHook _lowLevelMouseHook;
    private ClickCandidate? _lastClickCandidate;
    private ClickCandidate? _pendingNativeClose;
    private bool _suppressNextLeftUp;

    public ExplorerTabDoubleClickHook(ExplorerWatcher explorerWatcher)
    {
        _explorerWatcher = explorerWatcher;
        _lowLevelMouseHook = new LowLevelMouseHook
        {
            AddKeyboardKeys = true,
            Handling = true
        };
        _lowLevelMouseHook.Down += OnMouseDown;
        _lowLevelMouseHook.Up += OnMouseUp;
    }

    public event Action<string>? StatusChanged;
    public bool IsHookActive => _lowLevelMouseHook.IsStarted;

    public void StartHook() => _lowLevelMouseHook.Start();
    public void StopHook() => _lowLevelMouseHook.Stop();

    private void OnMouseDown(object? sender, MouseEventArgs e)
    {
        if (!IsLeftMouse(e))
            return;

        if (!SettingsManager.DoubleClickCloseTab)
        {
            ResetClickState();
            return;
        }

        var now = Environment.TickCount64;
        var currentPoint = e.Position;
        var previous = _lastClickCandidate;
        var explorerWindow = GetExplorerWindowForPoint(currentPoint);
        if (explorerWindow == 0)
        {
            if (previous == null ||
                !IsWithinDoubleClickWindow(previous, currentPoint, now) ||
                !Helper.IsFileExplorerWindow(previous.ExplorerWindow))
            {
                ResetClickState();
                return;
            }

            explorerWindow = previous.ExplorerWindow;
        }

        if (previous != null &&
            previous.ExplorerWindow == explorerWindow &&
            IsWithinDoubleClickWindow(previous, currentPoint, now) &&
            (_explorerWatcher.IsExplorerTabTitleAtScreenPoint(currentPoint, explorerWindow) ||
             _explorerWatcher.IsExplorerTabTitleAtScreenPoint(previous.Point, explorerWindow)))
        {
            e.IsHandled = true;
            _suppressNextLeftUp = true;
            _pendingNativeClose = new ClickCandidate(explorerWindow, currentPoint, now);
            _lastClickCandidate = null;
            return;
        }

        _lastClickCandidate = _explorerWatcher.IsExplorerTabTitleAtScreenPoint(currentPoint, explorerWindow)
            ? new ClickCandidate(explorerWindow, currentPoint, now)
            : null;
    }

    private void OnMouseUp(object? sender, MouseEventArgs e)
    {
        if (!IsLeftMouse(e) || !_suppressNextLeftUp)
            return;

        e.IsHandled = true;
        _suppressNextLeftUp = false;

        var pending = _pendingNativeClose;
        _pendingNativeClose = null;
        if (pending == null)
            return;

        ThreadPool.QueueUserWorkItem(_ =>
        {
            Thread.Sleep(160);
            if (_explorerWatcher.TryCloseTabAtScreenPoint(pending.Point, pending.ExplorerWindow))
                StatusChanged?.Invoke("Closed Explorer tab by title double-click.");
        });
    }

    private static bool IsLeftMouse(MouseEventArgs e) => e.CurrentKey is Key.MouseLeft or Key.LButton;

    private static nint GetExplorerWindowForPoint(Point point)
    {
        var hit = WinApi.WindowFromPoint(point);
        if (hit != 0)
        {
            var root = WinApi.GetAncestor(hit, GA_ROOT);
            if (Helper.IsFileExplorerWindow(root))
                return root;
        }

        return Helper.IsFileExplorerForeground(out var foreground) && foreground != 0 ? foreground : 0;
    }

    private static bool IsWithinDoubleClickWindow(ClickCandidate previous, Point currentPoint, long now)
    {
        var maxTime = Math.Max(200, WinApi.GetDoubleClickTime());
        var maxX = Math.Max(4, WinApi.GetSystemMetrics(SM_CXDOUBLECLK));
        var maxY = Math.Max(4, WinApi.GetSystemMetrics(SM_CYDOUBLECLK));

        return now - previous.Tick <= maxTime &&
               Math.Abs(currentPoint.X - previous.Point.X) <= maxX &&
               Math.Abs(currentPoint.Y - previous.Point.Y) <= maxY;
    }

    private void ResetClickState()
    {
        _lastClickCandidate = null;
        if (_suppressNextLeftUp && (WinApi.GetAsyncKeyState((int)VirtualKey.LeftButton) & 0x8000) == 0)
        {
            _suppressNextLeftUp = false;
            _pendingNativeClose = null;
        }
    }

    private sealed record ClickCandidate(nint ExplorerWindow, Point Point, long Tick);

    private enum VirtualKey
    {
        LeftButton = 0x01
    }

    public void Dispose()
    {
        StopHook();
        _lowLevelMouseHook.Dispose();
    }
}
