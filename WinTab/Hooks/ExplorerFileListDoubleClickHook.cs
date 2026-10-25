using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>The two native, same-tab navigation choices; neither creates a tab or synthesizes keys.</summary>
public enum FileListDoubleClickAction { Parent, Back }

/// <summary>
/// Observes native clicks without consuming input. Both backgrounds must be positively identified before
/// one short-lived navigation can run. Adapted from hbiui/WinTab 4121b08; slow UIA work is bounded and
/// unrelated input retires the gesture rather than allowing a late navigation on the user's next page.
/// </summary>
internal sealed class ExplorerFileListDoubleClickHook : IHook
{
    internal const int RequestLifetimeMs = 750;
    private readonly object _gate = new();
    private readonly Func<ExplorerFileListDoubleClickHook, IDisposable> _observeInput;
    private readonly Func<Point, long, ExplorerFileListClick?> _capture;
    private readonly Func<ExplorerFileListClick, Task<bool>> _hitTest;
    private readonly Func<ExplorerFileListDoubleClickRequest, FileListDoubleClickAction, Func<bool>, Task<bool>> _navigate;
    private readonly Func<ExplorerFileListDoubleClickRequest, bool> _isCurrent;
    private readonly Func<FileListDoubleClickAction> _action;
    private readonly Func<bool> _hasModifiers;
    private readonly ExplorerFileListDoubleClickState _state;
    private IDisposable? _input;
    private Task<bool>? _firstHit;
    private Task<bool>? _secondHit;
    private FileListDoubleClickAction _gestureAction;
    private nint _gestureWindow;
    private Point _gesturePoint;
    private int _generation;
    private int _hitWorkers;
    private volatile bool _active;
    private bool _disposed;

    internal ExplorerFileListDoubleClickHook(NavigationInputSource input, ExplorerWatcher watcher,
        Func<FileListDoubleClickAction> action)
        : this(hook => input.Subscribe(hook.ProcessPointer, hook.CancelPending, hook.OnForeground),
            (point, tick) => Capture(watcher, point, tick),
            click => Task.FromResult(ExplorerFileListAutomation.IsEmptyArea(click)),
            watcher.NavigateFileListAsync,
            request => ExplorerNavigationAccess.ForegroundFrame() == request.Second.Window.Handle &&
                ExplorerFileListAutomation.IsCurrent(request.First) && ExplorerFileListAutomation.IsCurrent(request.Second),
            action, ExplorerNavigationAccess.HasModifiers,
            new ExplorerFileListDoubleClickState(() => (int)WinApi.GetDoubleClickTime(),
                () => WinApi.GetSystemMetrics(36), () => WinApi.GetSystemMetrics(37))) { }

    // The same input and asynchronous pipeline is exercised by tests without native windows or input.
    internal ExplorerFileListDoubleClickHook(Func<ExplorerFileListDoubleClickHook, IDisposable> observeInput,
        Func<Point, long, ExplorerFileListClick?> capture, Func<ExplorerFileListClick, Task<bool>> hitTest,
        Func<ExplorerFileListDoubleClickRequest, FileListDoubleClickAction, Func<bool>, Task<bool>> navigate,
        Func<ExplorerFileListDoubleClickRequest, bool> isCurrent, Func<FileListDoubleClickAction> action,
        Func<bool> hasModifiers, ExplorerFileListDoubleClickState state)
    {
        _observeInput = observeInput;
        _capture = capture;
        _hitTest = hitTest;
        _navigate = navigate;
        _isCurrent = isCurrent;
        _action = action;
        _hasModifiers = hasModifiers;
        _state = state;
    }

    public bool IsHookActive => _active;

    public void StartHook()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_active) return;
        _active = true;
        try { _input = _observeInput(this); }
        catch { _active = false; throw; }
    }

    public void StopHook()
    {
        IDisposable? input;
        lock (_gate)
        {
            _active = false;
            CancelPendingLocked();
            input = _input;
            _input = null;
        }
        input?.Dispose();
    }

    private static ExplorerFileListClick? Capture(ExplorerWatcher watcher, Point point, long tick)
    {
        if (!ExplorerFileListAutomation.TryCaptureView(point, out var window, out var tab, out var view)) return null;
        var location = watcher.TryGetFileListLocationSnapshot(window, tab);
        if (location == null || !WinApi.GetWindowRect(view.Handle, out var bounds)) return null;
        return new ExplorerFileListClick(window, tab, view, point, location, tick) { ViewBounds = bounds };
    }

    /// <summary>The returned task represents this input's navigation, not any orphaned UIA provider call.</summary>
    internal Task ProcessPointerAsync(NavigationPointerInput input)
    {
        lock (_gate)
        {
            if (!_active) return Task.CompletedTask;
            if (input.FromWinTab || _hasModifiers() || input.DelayMs > 100 ||
                input.Kind is NavigationPointerKind.OtherDown or NavigationPointerKind.MiddleDown or NavigationPointerKind.Wheel)
            {
                CancelPendingLocked();
                return Task.CompletedTask;
            }
            if (input.Kind == NavigationPointerKind.Move)
            {
                // Include travel between clicks: moving away and back is not a double-click gesture.
                if (_gestureWindow != 0 && !_state.WithinDistance(_gesturePoint, input.Point)) CancelPendingLocked();
                return Task.CompletedTask;
            }
            var tick = Environment.TickCount64 - Math.Max(0, input.DelayMs);
            if (input.Kind == NavigationPointerKind.LeftDown)
            {
                var click = _capture(input.Point, tick);
                _generation++;
                var first = _state.HandleLeftDown(click);
                if (first != null)
                {
                    _firstHit = StartHit(first);
                    _secondHit = null;
                    _gestureWindow = first.Window.Handle;
                    _gesturePoint = first.Point;
                    _gestureAction = _action();
                }
                else if (_state.HasPair && click != null && _gestureAction == _action())
                    _secondHit = StartHit(click);
                else
                    CancelPendingLocked();
                return Task.CompletedTask;
            }
            if (input.Kind != NavigationPointerKind.LeftUp) return Task.CompletedTask;
            var request = _state.HandleLeftUp(input.Point, tick);
            if (request == null) return Task.CompletedTask;
            var firstHit = _firstHit;
            var secondHit = _secondHit;
            _firstHit = _secondHit = null;
            return firstHit != null && secondHit != null
                ? ProcessRequestAsync(request, firstHit, secondHit, _generation, _gestureAction)
                : Task.CompletedTask;
        }
    }

    private void ProcessPointer(NavigationPointerInput input) => _ = ProcessPointerAsync(input);

    private Task<bool> StartHit(ExplorerFileListClick click)
    {
        // UIA calls cannot be forcibly cancelled. Do not queue or create more workers when a provider hangs.
        if (Interlocked.Increment(ref _hitWorkers) > 2)
        {
            Interlocked.Decrement(ref _hitWorkers);
            return Task.FromResult(false);
        }
        return Task.Run(async () =>
        {
            try
            {
                if (Environment.TickCount64 - click.Tick > RequestLifetimeMs) return false;
                return await _hitTest(click).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                ExplorerDebugLog.Write($"File-list hit test failed: {exception.GetType().Name}:{exception.Message}");
                return false;
            }
            finally { Interlocked.Decrement(ref _hitWorkers); }
        });
    }

    private async Task ProcessRequestAsync(ExplorerFileListDoubleClickRequest request, Task<bool> firstHit,
        Task<bool> secondHit, int generation, FileListDoubleClickAction action)
    {
        bool Current() => _active && generation == Volatile.Read(ref _generation) && action == _action() &&
            Environment.TickCount64 - request.Second.Tick is >= 0 and <= RequestLifetimeMs && _isCurrent(request);
        try
        {
            var remaining = RequestLifetimeMs - (Environment.TickCount64 - request.Second.Tick);
            if (remaining <= 0 || !Current()) return;
            var hits = await Task.WhenAll(firstHit, secondHit).WaitAsync(TimeSpan.FromMilliseconds(remaining)).ConfigureAwait(false);
            if (hits[0] && hits[1] && Current())
                await _navigate(request, action, Current).ConfigureAwait(false);
        }
        catch (TimeoutException) { ExplorerDebugLog.Write("File-list double-click expired while waiting for hit tests"); }
        catch (Exception exception) { ExplorerDebugLog.Write($"File-list navigation failed: {exception.GetType().Name}:{exception.Message}"); }
    }

    private void OnForeground(nint window)
    {
        var root = WinApi.GetAncestor(window, WinApi.GA_ROOT);
        if (root != 0) window = root;
        lock (_gate)
            if (_gestureWindow != 0 && window != _gestureWindow) CancelPendingLocked();
    }

    internal void CancelPending()
    {
        lock (_gate) CancelPendingLocked();
    }

    private void CancelPendingLocked()
    {
        _generation++;
        _state.Reset();
        _firstHit = _secondHit = null;
        _gestureWindow = 0;
    }

    public void Dispose()
    {
        if (_disposed) return;
        StopHook();
        _disposed = true;
    }
}
