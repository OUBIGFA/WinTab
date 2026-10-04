using System;
using System.Collections.Generic;
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
    private readonly ExplorerWatcher _watcher;
    private readonly CoalescingAsyncWork _activationWork;
    private readonly SemaphoreSlim _activationGate = new(1);
    private readonly Dictionary<WindowIdentity, TabActivationHistory> _activationHistory = [];
    private readonly LowLevelMouseHook _lowLevelMouseHook;
    private readonly ExplorerTabDoubleClickCloseController _controller;
    private readonly Func<bool> _isEnabled;
    private readonly Func<bool> _includeNotepad;
    private readonly object _closeQueueGate = new();
    private Task _closeQueueTail = Task.CompletedTask;
    private int _generation;
    private int _clickSequence;
    private int _historyGeneration;
    private volatile bool _active;
    private bool _disposed;

    public ExplorerTabDoubleClickHook(ExplorerWatcher explorerWatcher, Func<bool>? isEnabled = null, Func<bool>? includeNotepad = null)
    {
        _tabStrip = explorerWatcher.TabStrip;
        _watcher = explorerWatcher;
        _activationWork = new CoalescingAsyncWork(ObserveActiveTabsAsync);
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
        _lowLevelMouseHook.Wheel += (_, _) => Interlocked.Increment(ref _clickSequence);
        _watcher.TabActivity += OnTabActivity;
    }

    public event Action<string>? StatusChanged;
    public bool IsHookActive => _lowLevelMouseHook.IsStarted;

    public void StartHook()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        _controller.Reset();
        _lowLevelMouseHook.Start();
        _active = true;
        _activationWork.Request();
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
        if (e.CurrentKey is Key.MouseRight or Key.RButton)
            Interlocked.Increment(ref _clickSequence);
        if (!IsLeftMouse(e))
            return;

        Interlocked.Increment(ref _clickSequence);

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
        var clickSequence = Volatile.Read(ref _clickSequence);
        bool IsLifetimeCurrent() => _active && _isEnabled() && generation == Volatile.Read(ref _generation) &&
            Environment.TickCount64 - queuedAt <= 500 && identity.IsCurrent &&
            ExplorerWindowDiscovery.IsDoubleClickCloseTarget(identity.Handle, _includeNotepad());
        bool IsCurrent() => IsLifetimeCurrent() &&
            WinApi.GetWindowRect(identity.Handle, out var rect) && rect.Equals(initialRect) &&
            WinApi.GetCursorPos(out var point) && point == closeRequest.Point &&
            GetDoubleClickTargetWindowForPoint(point, _includeNotepad()) == identity.Handle;
        lock (_closeQueueGate)
        {
            // Preserve the order of rapid double-clicks. Independent thread-pool work items can otherwise
            // issue tab closes out of order while Explorer is still rebuilding the tab strip.
            _closeQueueTail = _closeQueueTail
                .ContinueWith(_ => ExecuteCloseAsync(closeRequest, IsCurrent,
                    // Once the close is delivered, moving the pointer is harmless. Only a new action
                    // or leaving the window cancels returning to the previous tab.
                    () => IsLifetimeCurrent() && clickSequence == Volatile.Read(ref _clickSequence) &&
                        WinApi.GetForegroundWindow() == identity.Handle), CancellationToken.None,
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

    private async Task ExecuteCloseAsync(ExplorerTabCloseRequest closeRequest, Func<bool> isCurrent, Func<bool> canReturn)
    {
        await _activationGate.WaitAsync().ConfigureAwait(false);
        try
        {
            if (!isCurrent())
            {
                ExplorerDebugLog.Write("Double-click close expired before reading tabs");
                return;
            }
            var window = closeRequest.ExplorerWindow;
            var before = ExplorerTabAutomation.ReadSessionTabs(window);
            // The first click activates a background tab. Notepad can temporarily unpublish its strip
            // during that activation too; keep the same gesture until its live hit test is available.
            while (before.Length == 0 && isCurrent())
            {
                await Task.Delay(15).ConfigureAwait(false);
                before = ExplorerTabAutomation.ReadSessionTabs(window);
            }
            var closing = before.Where(tab => tab.Bounds.Contains(closeRequest.Point.X, closeRequest.Point.Y)).ToArray();
            if (closing.Length != 1 || !isCurrent())
            {
                ExplorerDebugLog.Write($"Double-click close no longer current after reading {before.Length} tabs hits={closing.Length}");
                return;
            }
            var history = GetHistory(WindowIdentity.Capture(window));
            history.Observe(before);
            var returnOrder = history.GetReturnOrder(closing[0].Id);
            bool delivered;
            if (returnOrder.Length > 0)
            {
                // Drain the native left-up before selecting: it must not reactivate the closing tab.
                await Task.Delay(10).ConfigureAwait(false);
                delivered = canReturn() && ExplorerTabAutomation.TryCloseTabReturningTo(window,
                    closing[0].Id, returnOrder[0], canReturn);
            }
            else
            {
                delivered = await ExecuteCloseWhenCurrentAsync(isCurrent,
                    () => MouseSimulator.SendMiddleClick(closeRequest.Point));
            }
            if (delivered)
            {
                // Only observe completion here. Selecting after the close would expose the native adjacent
                // tab first; the successor must have been selected before closing the original tab.
                await ObserveCloseAsync(window, before, closing[0].Id, returnOrder, canReturn);
                StatusChanged?.Invoke("Tab close requested");
            }
            else
                ExplorerDebugLog.Write($"Double-click close cancelled or exact tab action unavailable hwnd={window}");
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"Double-click close not delivered: {exception.GetType().Name}:{exception.Message}");
        }
        finally
        {
            _activationGate.Release();
            _activationWork.Request();
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

    private static async Task ObserveCloseAsync(nint window, ExplorerTabAutomation.Tab[] before,
        string closingId, string[] order, Func<bool> isCurrent)
    {
        if (order.Length == 0) return;
        while (isCurrent())
        {
            var after = ExplorerTabAutomation.ReadSessionTabs(window);
            if (!isCurrent()) return;
            if (after.Any(tab => tab.Id == closingId) || after.Length != before.Length - 1 ||
                after.Count(tab => tab.Selected) != 1)
            {
                // Both apps briefly unpublish their tab items while rebuilding the strip after a close.
                // An empty intermediate read is not confirmation that the surviving tabs disappeared.
                await Task.Delay(15).ConfigureAwait(false);
                continue;
            }
            var target = TabActivationHistory.FindReturnTab(order, closingId, before, after);
            ExplorerDebugLog.Write($"Double-click return hwnd={window} targetIndex={Array.FindIndex(after, tab => tab.Id == target)} selectedIndex={Array.FindIndex(after, tab => tab.Selected)} candidates={order.Length}");
            if (target != null && !after.Any(tab => tab.Id == target && tab.Selected))
                ExplorerDebugLog.Write($"Double-click return selection not observed hwnd={window}");
            return;
        }
    }

    private void OnTabActivity(nint window)
    {
        if (!_active || !_isEnabled()) return;
        var root = WinApi.GetAncestor(window, WinApi.GA_ROOT);
        if (root == WinApi.GetForegroundWindow() &&
            ExplorerWindowDiscovery.IsDoubleClickCloseTarget(root, _includeNotepad()))
            _activationWork.Request();
    }

    private async Task ObserveActiveTabsAsync()
    {
        var generation = Volatile.Read(ref _generation);
        await _activationGate.WaitAsync().ConfigureAwait(false);
        try
        {
            if (!_active || !_isEnabled() || generation != Volatile.Read(ref _generation)) return;
            foreach (var retired in _activationHistory.Keys.Where(window => !window.IsCurrent).ToArray())
                _activationHistory.Remove(retired);
            var window = WinApi.GetForegroundWindow();
            if (!ExplorerWindowDiscovery.IsDoubleClickCloseTarget(window, _includeNotepad())) return;
            var identity = WindowIdentity.Capture(window);
            var tabs = ExplorerTabAutomation.ReadSessionTabs(window);
            if (_active && generation == Volatile.Read(ref _generation) && identity.IsCurrent &&
                WinApi.GetForegroundWindow() == window)
                GetHistory(identity).Observe(tabs);
        }
        finally { _activationGate.Release(); }
    }

    private TabActivationHistory GetHistory(WindowIdentity window)
    {
        // Toggling the feature starts a fresh history; activations while it was off are unknown.
        var generation = Volatile.Read(ref _generation);
        if (_historyGeneration != generation)
        {
            _activationHistory.Clear();
            _historyGeneration = generation;
        }
        if (!_activationHistory.TryGetValue(window, out var history))
            _activationHistory.Add(window, history = new TabActivationHistory());
        return history;
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
        _watcher.TabActivity -= OnTabActivity;
        _ = _activationWork.StopAsync();
        _lowLevelMouseHook.Dispose();
    }
}
