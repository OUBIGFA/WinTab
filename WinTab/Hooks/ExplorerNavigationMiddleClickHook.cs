using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Leaves Explorer's middle click untouched and brings the tab it opens to the front. Nothing is prepared
/// ahead of a click: the button-down records the window's tabs, and the tab that appears afterwards is
/// selected with Explorer's native appended-tab command. There is no cache that can go stale between clicks.
/// </summary>
public sealed class ExplorerNavigationMiddleClickHook : IHook
{
    private readonly NavigationClickGate _clicks = new();
    private NavigationInputObserver? _input;
    private volatile bool _active;
    private bool _disposed;

    public event Action<string>? StatusChanged;
    public bool IsHookActive => _active;

    public void StartHook()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_active) return;
        _active = true;
        try
        {
            _input = new NavigationInputObserver(ProcessPointer, CancelForKeyboard, OnForeground);
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
        _clicks.Cancel();
        _input?.Dispose();
        _input = null;
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
        if (_clicks.CancelIfForegroundChanged(window))
            ExplorerDebugLog.Write($"Navigation cancelled by foreground={window}");
    }

    private void CancelForKeyboard()
    {
        if (_clicks.Cancel()) ExplorerDebugLog.Write("Navigation cancelled by keyboard input.");
    }

    internal void ProcessPointer(NavigationPointerInput input)
    {
        if (!_active || input.FromWinTab) return;
        if (input.Kind == NavigationPointerKind.Move)
        {
            _clicks.Move(input.Point, WinApi.GetSystemMetrics(WinApi.SM_CXDRAG), WinApi.GetSystemMetrics(WinApi.SM_CYDRAG));
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
            if (_clicks.Cancel()) ExplorerDebugLog.Write($"Navigation cancelled by pointer={input.Kind}");
            return;
        }
        if (input.Kind != NavigationPointerKind.MiddleDown) return;
        if (HasModifiers()) { _clicks.Cancel(); return; }

        // Only window handles are read here; nothing on the native input path waits for Explorer's UI thread.
        var startedAt = Stopwatch.GetTimestamp();
        var window = WindowAt(input.Point);
        var tree = WinApi.WindowFromPoint(input.Point);
        if (window == 0 || !ExplorerNavigationAccess.IsNavigationTree(tree, window))
        {
            _clicks.Cancel();
            return;
        }
        var tabs = ExplorerNavigationAccess.Tabs(window);
        var sourceTab = ExplorerNavigationAccess.ActiveTab(window);
        if (tabs.Length is 0 or > ExplorerNavigationAccess.MaxTabs || sourceTab == 0)
        {
            _clicks.Cancel();
            ExplorerDebugLog.Write($"Navigation middle-click left native: tabs={tabs.Length} hwnd={window}");
            return;
        }
        var lease = _clicks.Begin(window, input.Point, Environment.TickCount64);
        if (lease == null) return;
        var click = new NavigationClick(WindowIdentity.Capture(window), WindowIdentity.Capture(sourceTab),
            WindowIdentity.Capture(tree), tabs, input.Point);
        ExplorerDebugLog.Write($"Navigation click captured hwnd={window} tab={sourceTab} tabs={tabs.Length} hookMs={Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F1}");
        _ = Task.Run(() => ActivateAsync(click, lease));
    }

    private async Task ActivateAsync(NavigationClick click, NavigationClickLease lease)
    {
        var result = NavigationActivationResult.Cancelled;
        var forgotten = false;
        var startedAt = Stopwatch.GetTimestamp();
        string Elapsed() => $"{Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F0}ms";
        bool Current()
        {
            var current = _active && _clicks.IsCurrent(lease, Environment.TickCount64) && click.IsCurrent();
            if (!current)
                ExplorerDebugLog.Write($"Navigation request retired active={_active} lease={_clicks.IsCurrent(lease, Environment.TickCount64)} window={click.Window.IsCurrent} source={click.SourceTab.IsCurrent} tree={click.Tree.IsCurrent} foreground={WinApi.GetForegroundWindow()}");
            return current;
        }
        try
        {
            // A click beside the folder items opens nothing, so it is forgotten at once and does not hold up the next click.
            var onFolder = await Task.Run(() => ExplorerNavigationAccess.IsFolderItemAt(click.Tree.Handle, click.Point))
                .WaitAsync(lease.Token).ConfigureAwait(false);
            if (!onFolder)
            {
                forgotten = true;
                _clicks.Discard(lease);
                ExplorerDebugLog.Write($"Navigation middle-click left native: not on a folder item hwnd={click.Window.Handle}");
                return;
            }
            ExplorerDebugLog.Write($"Navigation folder item confirmed at {Elapsed()}");
            await lease.Released.Task.WaitAsync(lease.Token).ConfigureAwait(false);
            ExplorerDebugLog.Write($"Navigation button released at {Elapsed()}");
            if (Current())
            {
                result = await NavigationTabActivation.RunAsync(click.TabsBefore, click.SourceTab.Handle,
                    Current, () => ExplorerNavigationAccess.Observe(click.Window.Handle),
                    newTab =>
                    {
                        ExplorerDebugLog.Write($"Navigation new tab={newTab} seen at {Elapsed()}");
                        var outcome = ExplorerNavigationAccess.SelectNewTab(click, newTab, Current);
                        ExplorerDebugLog.Write($"Navigation select outcome={outcome} at {Elapsed()}");
                        return outcome;
                    }, lease.Token,
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
            if (!forgotten)
                ExplorerDebugLog.Write($"Navigation activation result={result} hwnd={click.Window.Handle} at {Elapsed()}");
            _clicks.Complete(lease, result is NavigationActivationResult.Activated or NavigationActivationResult.AlreadyActive);
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
        GC.SuppressFinalize(this);
    }
}
