using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Leaves Explorer's middle click untouched and brings the tab it opens to the front, wherever inside the
/// active tab the folder was clicked: navigation pane, file list, Home page or address bar. Nothing is
/// prepared ahead of a click: the button-down records the window's tabs, and the tab that appears afterwards
/// is selected with Explorer's native appended-tab command. There is no cache that can go stale between clicks.
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
        if (hit == 0) return 0;

        // Address-bar XAML islands can make the root returned by GA_ROOT differ from the CabinetWClass
        // frame. Walk the parent chain as well so the click is still associated with the visible Explorer.
        var root = WinApi.GetAncestor(hit, WinApi.GA_ROOT);
        if (ExplorerWindowDiscovery.IsShownExplorerWindow(root))
            return root;
        nint current = hit;
        for (var depth = 0; current != 0 && depth < 64; depth++)
        {
            if (ExplorerWindowDiscovery.IsShownExplorerWindow(current))
                return current;
            current = WinApi.GetParent(current);
        }
        return 0;
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

    private void CancelForKeyboard() => CancelPending("keyboard input");

    private void CancelPending(string reason)
    {
        if (_clicks.Cancel()) ExplorerDebugLog.Write($"Navigation cancelled by {reason}");
    }

    internal void ProcessPointer(NavigationPointerInput input)
    {
        if (!_active || input.FromWinTab) return;
        if (input.Kind == NavigationPointerKind.MiddleUp)
        {
            if (HasModifiers()) CancelPending("a modifier key at button-up");
            else _clicks.Release(Environment.TickCount64);
            return;
        }
        if (input.Kind == NavigationPointerKind.OtherDown)
        {
            CancelPending($"pointer={input.Kind}");
            return;
        }
        if (HasModifiers()) { CancelPending("a middle click with a modifier key"); return; }

        // Only window handles are read here; nothing on the native input path waits for Explorer's UI thread.
        var startedAt = Stopwatch.GetTimestamp();
        var window = WindowAt(input.Point);
        var target = WinApi.WindowFromPoint(input.Point);
        if (window == 0 || (!ExplorerNavigationAccess.IsInsideActiveTab(target, window) &&
                            !ExplorerNavigationAccess.IsPotentialNavigationTarget(target, window)))
        {
            CancelPending("a middle click outside Explorer's tabs");
            return;
        }
        var onTree = ExplorerNavigationAccess.IsNavigationTree(target, window);
        var tabs = ExplorerNavigationAccess.ReadTabs(window);
        var sourceTab = tabs is { Length: > 0 } ? tabs[0] : 0;
        if (tabs is null || tabs.Length > ExplorerNavigationAccess.MaxTabs || sourceTab == 0)
        {
            CancelPending("a middle click in a window whose tabs cannot be followed");
            ExplorerDebugLog.Write($"Navigation middle-click left native: tabs={tabs?.Length.ToString() ?? "changing"} hwnd={window}");
            return;
        }
        var lease = _clicks.Begin(window, Environment.TickCount64, out var replaced);
        if (lease == null) return;
        var click = new NavigationClick(WindowIdentity.Capture(window), WindowIdentity.Capture(sourceTab),
            WindowIdentity.Capture(target), tabs, input.Point, onTree);
        ExplorerDebugLog.Write($"Navigation click captured hwnd={window} tab={sourceTab} target={target} tree={onTree} tabs={tabs.Length} replacedEarlier={replaced} hookMs={Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F1}");
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
                ExplorerDebugLog.Write($"Navigation request retired active={_active} lease={_clicks.IsCurrent(lease, Environment.TickCount64)} window={click.Window.IsCurrent} source={click.SourceTab.IsCurrent} target={click.Target.IsCurrent} foreground={WinApi.GetForegroundWindow()}");
            return current;
        }
        try
        {
            // On the navigation tree a click beside the folder items is known to open nothing, so it is forgotten
            // at once. Elsewhere (file list, Home page, address bar) the only evidence is the tab Explorer does or
            // does not open, so the click waits for it; a later click replaces it rather than waiting behind it.
            if (click.OnNavigationTree)
            {
                var onFolder = await Task.Run(() => ExplorerNavigationAccess.IsFolderItemAt(click.Target.Handle, click.Point))
                    .WaitAsync(lease.Token).ConfigureAwait(false);
                if (!onFolder)
                {
                    forgotten = true;
                    ExplorerDebugLog.Write($"Navigation middle-click left native: not on a folder item hwnd={click.Window.Handle}");
                    return;
                }
                ExplorerDebugLog.Write($"Navigation folder item confirmed at {Elapsed()}");
            }
            await lease.Released.Task.WaitAsync(lease.Token).ConfigureAwait(false);
            ExplorerDebugLog.Write($"Navigation button released at {Elapsed()}");
            if (Current())
            {
                result = await NavigationTabActivation.RunAsync(click.TabsBefore, click.SourceTab.Handle,
                    Current, () =>
                    {
                        var observation = ExplorerNavigationAccess.Observe(click.Window.Handle);
                        if (observation == null)
                            ExplorerDebugLog.Write($"Navigation tab windows moving; observing again at {Elapsed()}");
                        return observation;
                    },
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
            _clicks.Complete(lease);
        }
        if (result == NavigationActivationResult.Activated)
            StatusChanged?.Invoke("Activated Explorer tab opened by middle-click.");
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
