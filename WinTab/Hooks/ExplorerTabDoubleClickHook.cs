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
    // Bound stale requests, not the interval between gestures. There is no cooldown after delivery.
    private const int CloseRequestLifetimeMs = 500;

    private readonly TabStripHitTester _tabStrip;
    private readonly ExplorerWatcher _watcher;
    private readonly CoalescingAsyncWork _activationWork;
    private readonly object _historyGate = new();
    private readonly Dictionary<WindowIdentity, TabActivationHistory> _activationHistory = [];
    private readonly LowLevelMouseHook _lowLevelMouseHook;
    private readonly ExplorerTabDoubleClickCloseController _controller;
    private readonly ExplorerTabCloseQueue _closeQueue;
    private readonly Func<bool> _isEnabled;
    private readonly Func<bool> _includeNotepad;
    private int _generation;
    private int _closeRevision;
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
        _closeQueue = new ExplorerTabCloseQueue(new CloseEnvironment(this));
        _lowLevelMouseHook = new LowLevelMouseHook
        {
            AddKeyboardKeys = true,
            Handling = true,
            GenerateMouseMoveEvents = true
        };
        _lowLevelMouseHook.Down += OnMouseDown;
        _lowLevelMouseHook.Up += OnMouseUp;
        _lowLevelMouseHook.Move += (_, e) => _controller.HandleMouseMove(e.Position);
        _lowLevelMouseHook.Wheel += (_, _) => _controller.CancelGesture();
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
        if (!IsLeftMouse(e))
        {
            _controller.CancelGesture();
            return;
        }

        var decision = _controller.HandleLeftMouseDown(e.Position, Environment.TickCount64);
        if (decision.Handled)
            e.IsHandled = true;
    }

    private void OnMouseUp(object? sender, MouseEventArgs e)
    {
        if (!IsLeftMouse(e))
            return;

        var decision = _controller.HandleLeftMouseUp(Environment.TickCount64, e.Position);
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
        var interaction = _controller.InteractionVersion;
        bool CanAct() => _active && _isEnabled() && generation == Volatile.Read(ref _generation) &&
            Environment.TickCount64 - queuedAt <= CloseRequestLifetimeMs && identity.IsCurrent &&
            interaction == _controller.InteractionVersion &&
            ExplorerWindowDiscovery.IsDoubleClickCloseTarget(identity.Handle, _includeNotepad()) &&
            WinApi.GetForegroundWindow() == identity.Handle &&
            WinApi.GetWindowRect(identity.Handle, out var rect) && rect.Equals(initialRect) &&
            GetDoubleClickTargetWindowForPoint(closeRequest.Point, _includeNotepad()) == identity.Handle;
        bool IsCurrent() => CanAct() && WinApi.GetCursorPos(out var point) &&
            ExplorerTabDoubleClickCloseController.IsWithinDoubleClickDistance(closeRequest.Point, point,
                WinApi.GetSystemMetrics(SM_CXDOUBLECLK), WinApi.GetSystemMetrics(SM_CYDOUBLECLK));

        Interlocked.Increment(ref _closeRevision);
        _ = ReportCloseAsync(closeRequest, _closeQueue.QueueAsync(closeRequest, IsCurrent, CanAct), queuedAt);
    }

    private async Task ReportCloseAsync(ExplorerTabCloseRequest request, Task<bool> work, long queuedAt)
    {
        try
        {
            if (await work.ConfigureAwait(false))
                StatusChanged?.Invoke("Tab close requested");
            else
                ExplorerDebugLog.Write($"Double-click close cancelled or exact tab unavailable hwnd={request.ExplorerWindow} elapsed={Environment.TickCount64 - queuedAt}ms");
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"Double-click close not delivered: {exception.GetType().Name}:{exception.Message}");
        }
        finally
        {
            _activationWork.Request();
            // Keep the last measured geometry usable until the new snapshot arrives. Clearing it after
            // every close created a blind interval for the next pair, especially at a different position.
            if (ExplorerWindowDiscovery.IsFileExplorerWindow(request.ExplorerWindow))
                _tabStrip.ScheduleRefresh(request.ExplorerWindow);
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

    private Task ObserveActiveTabsAsync()
    {
        if (!_active || !_isEnabled() || _closeQueue.HasPending) return Task.CompletedTask;
        var generation = Volatile.Read(ref _generation);
        var revision = Volatile.Read(ref _closeRevision);
        var window = WinApi.GetForegroundWindow();
        if (!ExplorerWindowDiscovery.IsDoubleClickCloseTarget(window, _includeNotepad())) return Task.CompletedTask;
        var identity = WindowIdentity.Capture(window);
        // Never hold the history lock across UIA. A slow passive scan must not use up a close's lifetime.
        var tabs = ExplorerTabAutomation.ReadSessionTabs(window);
        lock (_historyGate)
        {
            if (!_active || generation != Volatile.Read(ref _generation) || !identity.IsCurrent ||
                WinApi.GetForegroundWindow() != window || _closeQueue.HasPending ||
                revision != Volatile.Read(ref _closeRevision)) return Task.CompletedTask;
            foreach (var retired in _activationHistory.Keys.Where(key => !key.IsCurrent).ToArray())
                _activationHistory.Remove(retired);
            GetHistory(identity).Observe(tabs);
        }
        return Task.CompletedTask;
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
        public bool? IsPointOnTabStrip(Point point, nint explorerWindow)
        {
            if (ExplorerWindowDiscovery.IsNotepadWindow(explorerWindow)) return null;
            var hit = tabStrip.HitTestTab(point, explorerWindow);
            // A tab may expand into previously empty title-row space during close/selection layout.
            // Revalidate such a hit off the hook rather than discarding the first pair at its new position.
            return hit == false && tabStrip.TryGetTabRow(explorerWindow, out var row) && row.Contains(point)
                ? null : hit;
        }
    }

    private sealed class CloseEnvironment(ExplorerTabDoubleClickHook owner) : IExplorerTabCloseEnvironment
    {
        public ExplorerTabAutomation.Tab[] ReadTabs(nint window) => ExplorerTabAutomation.ReadSessionTabs(window);

        public string? GetReturnTab(nint window, ExplorerTabAutomation.Tab[] tabs, string closingId)
        {
            lock (owner._historyGate)
            {
                var history = owner.GetHistory(WindowIdentity.Capture(window));
                history.Observe(tabs);
                return history.GetReturnOrder(closingId).FirstOrDefault();
            }
        }

        public bool TryCloseTab(nint window, string closingId, string? returnId, Func<bool> isCurrent) =>
            ExplorerTabAutomation.TryCloseTabReturningTo(window, closingId, returnId, isCurrent);
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
