using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using H.Hooks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

public sealed class ExplorerTabDoubleClickHook : IHook
{
    private const int SM_CXDOUBLECLK = 36;
    private const int SM_CYDOUBLECLK = 37;

    private readonly TabStripHitTester _tabStrip;
    private readonly LowLevelMouseHook _lowLevelMouseHook;
    private readonly ExplorerTabDoubleClickCloseController _controller;
    private readonly Func<bool> _isEnabled;
    private readonly Func<bool> _includeNotepad;
    private readonly object _closeQueueGate = new();
    private Task _closeQueueTail = Task.CompletedTask;
    private int _generation;
    private volatile bool _active;
    private bool _disposed;

    public ExplorerTabDoubleClickHook(ExplorerWatcher explorerWatcher, Func<bool>? isEnabled = null, Func<bool>? includeNotepad = null)
    {
        _tabStrip = explorerWatcher.TabStrip;
        _isEnabled = isEnabled ?? (() => true);
        _includeNotepad = includeNotepad ?? (() => true);
        _controller = new ExplorerTabDoubleClickCloseController(new HookEnvironment(_tabStrip, () => _active && _isEnabled(), _includeNotepad));
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

    public void StartHook()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        _controller.Reset();
        _lowLevelMouseHook.Start();
        _active = true;
    }

    public void StopHook()
    {
        _active = false;
        Interlocked.Increment(ref _generation);
        _lowLevelMouseHook.Stop();
        _controller.Reset();
    }

    private void OnMouseDown(object? sender, MouseEventArgs e)
    {
        if (!IsLeftMouse(e))
            return;

        var decision = _controller.HandleLeftMouseDown(e.Position, Environment.TickCount64);
        if (decision.Handled)
            e.IsHandled = true;
    }

    private void OnMouseUp(object? sender, MouseEventArgs e)
    {
        if (!IsLeftMouse(e))
            return;

        var decision = _controller.HandleLeftMouseUp(Environment.TickCount64);
        if (decision.Handled)
            e.IsHandled = true;

        if (decision.CloseRequest is not { } closeRequest)
            return;

        QueueClose(closeRequest);
    }

    private void QueueClose(ExplorerTabCloseRequest closeRequest)
    {
        var identity = WindowIdentity.Capture(closeRequest.ExplorerWindow);
        if (!identity.IsCurrent || !WinApi.GetWindowRect(identity.Handle, out var initialRect)) return;
        var generation = Volatile.Read(ref _generation);
        var queuedAt = Environment.TickCount64;
        bool IsCurrent() => _active && _isEnabled() && generation == Volatile.Read(ref _generation) &&
            Environment.TickCount64 - queuedAt <= 500 && identity.IsCurrent &&
            WinApi.GetWindowRect(identity.Handle, out var rect) && rect.Equals(initialRect) &&
            WinApi.GetCursorPos(out var point) && point == closeRequest.Point &&
            GetDoubleClickTargetWindowForPoint(point, _includeNotepad()) == identity.Handle;
        lock (_closeQueueGate)
        {
            // Preserve the order of rapid double-clicks. Independent thread-pool work items can otherwise
            // inject middle-clicks out of order while Explorer is still rebuilding the tab strip.
            _closeQueueTail = _closeQueueTail
                .ContinueWith(_ => ExecuteCloseAsync(closeRequest, IsCurrent), CancellationToken.None,
                    TaskContinuationOptions.None, TaskScheduler.Default)
                .Unwrap();
        }
    }

    internal static async Task<bool> ExecuteCloseWhenCurrentAsync(Func<bool> isCurrent, Action close)
    {
        // Let the suppressed left-up drain, then recheck ownership immediately before synthesizing input.
        await Task.Delay(10).ConfigureAwait(false);
        if (!isCurrent()) return false;
        close();
        return true;
    }

    private async Task ExecuteCloseAsync(ExplorerTabCloseRequest closeRequest, Func<bool> isCurrent)
    {
        try
        {
            // Read Notepad's current tab rectangles once, off the hook. No cold-cache misses and no
            // close-button UIA walk. Recheck ownership after the remote read before injecting input.
            if (ExplorerWindowDiscovery.IsNotepadWindow(closeRequest.ExplorerWindow) &&
                (!isCurrent() || !ExplorerTabAutomation.ReadTabs(closeRequest.ExplorerWindow)
                    .Any(tab => tab.Bounds.Contains(closeRequest.Point.X, closeRequest.Point.Y))))
                return;
            var delivered = await ExecuteCloseWhenCurrentAsync(isCurrent,
                () => MouseSimulator.SendMiddleClick(closeRequest.Point));
            if (delivered)
                StatusChanged?.Invoke("Tab close requested");
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"Double-click close not delivered: {exception.GetType().Name}:{exception.Message}");
        }
        finally
        {
            try
            {
                if (ExplorerWindowDiscovery.IsFileExplorerWindow(closeRequest.ExplorerWindow))
                    _tabStrip.Refresh(closeRequest.ExplorerWindow);
            }
            catch
            {
                // 忽略刷新失败，后续鼠标事件会重新计算。
            }
        }
    }

    private static bool IsLeftMouse(MouseEventArgs e) => e.CurrentKey is Key.MouseLeft or Key.LButton;

    private static nint GetDoubleClickTargetWindowForPoint(Point point, bool includeNotepad)
    {
        var hit = WinApi.WindowFromPoint(point);
        if (hit != 0)
        {
            var root = WinApi.GetAncestor(hit, WinApi.GA_ROOT);
            if (ExplorerWindowDiscovery.IsDoubleClickCloseTarget(root, includeNotepad))
                return root;
        }

        return 0;
    }

    private sealed class HookEnvironment(TabStripHitTester tabStrip, Func<bool> isEnabled, Func<bool> includeNotepad) : IExplorerTabDoubleClickEnvironment
    {
        public bool IsEnabled => isEnabled();
        public int DoubleClickTimeMs => (int)WinApi.GetDoubleClickTime();
        public int DoubleClickWidth => WinApi.GetSystemMetrics(SM_CXDOUBLECLK);
        public int DoubleClickHeight => WinApi.GetSystemMetrics(SM_CYDOUBLECLK);
        public nint ResolveExplorerWindow(Point point) => GetDoubleClickTargetWindowForPoint(point, includeNotepad());
        public bool IsExplorerWindow(nint explorerWindow) => ExplorerWindowDiscovery.IsDoubleClickCloseTarget(explorerWindow, includeNotepad());
        public bool IsPointOnTabStrip(Point point, nint explorerWindow) => tabStrip.IsPointOnTabStrip(point, explorerWindow);
        public bool ShouldDeferHitTest(nint window) => ExplorerWindowDiscovery.IsNotepadWindow(window);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        StopHook();
        _lowLevelMouseHook.Dispose();
    }
}
