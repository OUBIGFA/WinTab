using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>Leaves Explorer's middle-click untouched and selects only a uniquely correlated new tab.</summary>
public sealed class ExplorerNavigationMiddleClickHook : IHook
{
    private readonly BackgroundRefreshCache<nint, ExplorerNavigationSnapshot> _snapshots = new(CaptureSnapshot);
    private readonly NavigationClickGate _clicks = new();
    private readonly Timer _warmTimer;
    private NavigationInputObserver? _input;
    private volatile bool _active;
    private bool _disposed;
    private nint _pendingWindow;
    private long _lastHoverWarm;

    public event Action<string>? StatusChanged;
    public bool IsHookActive => _active;

    public ExplorerNavigationMiddleClickHook() =>
        _warmTimer = new Timer(_ => WarmForeground(), null, Timeout.Infinite, Timeout.Infinite);

    public void StartHook()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_active) return;
        _active = true;
        try
        {
            _input = new NavigationInputObserver(ProcessPointer, CancelForKeyboard, OnForeground);
            _warmTimer.Change(0, 200);
        }
        catch
        {
            _active = false;
            throw;
        }
    }

    public void StopHook()
    {
        _active = false;
        _warmTimer.Change(Timeout.Infinite, Timeout.Infinite);
        _clicks.Cancel();
        _input?.Dispose();
        _input = null;
        _snapshots.Clear();
    }

    private static ExplorerNavigationSnapshot? CaptureSnapshot(nint window)
    {
        try { return ExplorerNavigationAccess.Capture(window); }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"Navigation snapshot unavailable hwnd={window} error={exception.GetType().Name}:{exception.Message}");
            return null;
        }
    }

    internal bool IsPrepared(nint window) =>
        _snapshots.TryGet(window, out var snapshot) && snapshot!.Matches(window,
            ExplorerNavigationAccess.NavigationTree(ExplorerNavigationAccess.ActiveTab(window)), ExplorerNavigationAccess.Tabs(window));

    private void WarmForeground()
    {
        if (_active) Warm(WinApi.GetForegroundWindow());
    }

    private void Warm(nint window)
    {
        if (!_active || !ExplorerWindowDiscovery.IsShownExplorerWindow(window)) return;
        var tree = ExplorerNavigationAccess.NavigationTree(ExplorerNavigationAccess.ActiveTab(window));
        if (tree == 0) return;
        if (!_snapshots.TryGet(window, out var snapshot) ||
            Environment.TickCount64 - snapshot!.CreatedAt > 1_000 ||
            !snapshot.Matches(window, tree, ExplorerNavigationAccess.Tabs(window)))
            _snapshots.Request(window);
    }

    private static nint WindowAt(Point point)
    {
        var hit = WinApi.WindowFromPoint(point);
        var root = hit == 0 ? 0 : WinApi.GetAncestor(hit, WinApi.GA_ROOT);
        return ExplorerWindowDiscovery.IsShownExplorerWindow(root) ? root : 0;
    }

    private void OnForeground(nint window)
    {
        // XAML tab selection can report its InputSiteWindowClass child as the foreground event HWND.
        // That is still the same Explorer frame, not the user switching to another application.
        var root = WinApi.GetAncestor(window, WinApi.GA_ROOT);
        if (root != 0) window = root;
        if (window != _pendingWindow)
        {
            if (_pendingWindow != 0) ExplorerDebugLog.Write($"Navigation cancelled by foreground={window}");
            _clicks.Cancel();
        }
        if (_active && ExplorerWindowDiscovery.IsFileExplorerWindow(window)) _snapshots.Request(window);
    }

    private void CancelForKeyboard()
    {
        if (_pendingWindow != 0) ExplorerDebugLog.Write("Navigation cancelled by keyboard input.");
        _clicks.Cancel();
        var window = WinApi.GetForegroundWindow();
        if (_active && ExplorerWindowDiscovery.IsFileExplorerWindow(window)) _snapshots.Invalidate(window);
    }

    internal void ProcessPointer(NavigationPointerInput input)
    {
        if (!_active || input.Injected) return;
        if (input.Kind == NavigationPointerKind.Move)
        {
            _clicks.Move(input.Point, WinApi.GetSystemMetrics(WinApi.SM_CXDRAG), WinApi.GetSystemMetrics(WinApi.SM_CYDRAG));
            var now = Environment.TickCount64;
            if (now - Interlocked.Read(ref _lastHoverWarm) >= 100)
            {
                Interlocked.Exchange(ref _lastHoverWarm, now);
                // Only enqueue here; accessibility is never queried by a hook callback.
                var hovered = WindowAt(input.Point);
                if (hovered != 0 && !_snapshots.TryGet(hovered, out _)) _snapshots.Request(hovered);
            }
            return;
        }
        if (input.Kind == NavigationPointerKind.MiddleUp)
        {
            if (HasModifiers()) _clicks.Cancel();
            else _clicks.Release(input.Point, Environment.TickCount64, WinApi.GetSystemMetrics(WinApi.SM_CXDRAG), WinApi.GetSystemMetrics(WinApi.SM_CYDRAG));
            return;
        }
        if (input.Kind is NavigationPointerKind.OtherDown or NavigationPointerKind.Wheel)
        {
            if (_pendingWindow != 0) ExplorerDebugLog.Write($"Navigation cancelled by pointer={input.Kind}");
            _clicks.Cancel();
            var changedWindow = WindowAt(input.Point);
            if (changedWindow != 0) _snapshots.Invalidate(changedWindow);
            return;
        }
        if (input.Kind != NavigationPointerKind.MiddleDown) return;
        if (HasModifiers()) { _clicks.Cancel(); return; }

        var startedAt = Stopwatch.GetTimestamp();
        var window = WindowAt(input.Point);
        var tree = WinApi.WindowFromPoint(input.Point);
        var tabs = window == 0 ? [] : ExplorerNavigationAccess.Tabs(window);
        if (window == 0 || tabs.Length is 0 or > ExplorerNavigationAccess.MaxTabs ||
            !WinApi.IsWindowHasClassName(tree, "SysTreeView32") ||
            !_snapshots.TryGet(window, out var snapshot) || !snapshot!.Matches(window, tree, tabs))
        {
            _clicks.Cancel();
            if (window != 0) _snapshots.Request(window);
            ExplorerDebugLog.Write("Navigation middle-click left native: no matching prepared snapshot.");
            return;
        }
        var item = Array.Find(snapshot.Items, row => row.HitBounds.Contains(input.Point));
        // Cap synchronous work. Nothing on the native input path waits for another process's UI thread.
        if (item == null || Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds > 12)
        {
            _clicks.Cancel();
            return;
        }
        var lease = _clicks.Begin(input.Point, Environment.TickCount64);
        if (lease == null) return;
        _pendingWindow = window;
        ExplorerDebugLog.Write($"Navigation click captured hwnd={window} tab={snapshot.SourceTab.Handle} item={item.Id} hookMs={Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F1}");
        _ = Task.Run(() => ActivateAsync(snapshot, item, lease));
    }

    private async Task ActivateAsync(ExplorerNavigationSnapshot snapshot, NavigationItem item, NavigationClickLease lease)
    {
        var result = NavigationActivationResult.Cancelled;
        bool Current()
        {
            var current = _active && _clicks.IsCurrent(lease, Environment.TickCount64) &&
                snapshot.Window.IsCurrent && snapshot.SourceTab.IsCurrent && snapshot.Tree.IsCurrent &&
                WinApi.GetParent(snapshot.SourceTab.Handle) == snapshot.Window.Handle &&
                WinApi.GetForegroundWindow() == snapshot.Window.Handle && snapshot.HasSameBounds();
            if (!current)
                ExplorerDebugLog.Write($"Navigation request retired active={_active} lease={_clicks.IsCurrent(lease, Environment.TickCount64)} window={snapshot.Window.IsCurrent} source={snapshot.SourceTab.IsCurrent} tree={snapshot.Tree.IsCurrent} parent={WinApi.GetParent(snapshot.SourceTab.Handle)} foreground={WinApi.GetForegroundWindow()} sameBounds={snapshot.HasSameBounds()}");
            return current;
        }
        try
        {
            await lease.Released.Task.WaitAsync(lease.Token).ConfigureAwait(false);
            if (Current() && ExplorerNavigationAccess.HitStillMatches(snapshot, item, lease.Point))
            {
                result = await NavigationTabActivation.RunAsync(snapshot.Tabs, snapshot.SourceTab.Handle,
                    Current, () => ExplorerNavigationAccess.Observe(snapshot.Window.Handle, snapshot.Tabs),
                    id => ExplorerNavigationAccess.Select(snapshot, item, lease.Point, id, Current), lease.Token,
                    Math.Max(1, (int)(lease.Deadline - Environment.TickCount64))).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception exception)
        {
            result = NavigationActivationResult.SelectionRejected;
            ExplorerDebugLog.Write($"Navigation activation failed: {exception.GetType().Name}:{exception.Message}");
        }
        finally
        {
            ExplorerDebugLog.Write($"Navigation activation result={result} hwnd={snapshot.Window.Handle}");
            _pendingWindow = 0;
            _clicks.Complete(lease, result is NavigationActivationResult.Activated or NavigationActivationResult.AlreadyActive);
            if (_active) _snapshots.Invalidate(snapshot.Window.Handle);
        }
        if (result == NavigationActivationResult.Activated)
            StatusChanged?.Invoke("Activated Explorer tab opened by navigation-pane middle-click.");
    }

    private static bool HasModifiers() =>
        (WinApi.GetAsyncKeyState(WinApi.VK_SHIFT) & 0x8000) != 0 ||
        (WinApi.GetAsyncKeyState(WinApi.VK_CONTROL) & 0x8000) != 0 ||
        (WinApi.GetAsyncKeyState(WinApi.VK_MENU) & 0x8000) != 0 ||
        (WinApi.GetAsyncKeyState(WinApi.VK_LWIN) & 0x8000) != 0 ||
        (WinApi.GetAsyncKeyState(WinApi.VK_RWIN) & 0x8000) != 0;

    public void Dispose()
    {
        if (_disposed) return;
        StopHook();
        _disposed = true;
        _clicks.Dispose();
        _warmTimer.Dispose();
        _snapshots.Dispose();
        GC.SuppressFinalize(this);
    }
}
