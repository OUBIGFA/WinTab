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
    private readonly Func<ExplorerNavigationMiddleClickHook, IDisposable> _observeInput;
    private IDisposable? _input;
    private volatile bool _active;
    private bool _disposed;

    public ExplorerNavigationMiddleClickHook()
        : this(hook => new NavigationInputObserver(hook.ProcessPointer, hook.CancelForKeyboard, hook.OnForeground))
    {
    }

    /// <param name="observeInput">Starts delivering the user's input to the hook; tests deliver it themselves.</param>
    internal ExplorerNavigationMiddleClickHook(Func<ExplorerNavigationMiddleClickHook, IDisposable> observeInput) =>
        _observeInput = observeInput;

    public event Action<string>? StatusChanged;
    public bool IsHookActive => _active;

    public void StartHook()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_active) return;
        _active = true;
        try
        {
            _input = _observeInput(this);
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
        var owner = WinApi.GetAncestor(window, WinApi.GA_ROOTOWNER);
        if (_clicks.CancelIfForegroundChanged(window, owner))
            ExplorerDebugLog.Write($"Navigation cancelled by foreground={window}:{WinApi.GetWindowClassName(window)} owner={owner}");
    }

    private void CancelForKeyboard() => CancelPending("keyboard input");

    private void CancelPending(string reason)
    {
        if (_clicks.Cancel()) ExplorerDebugLog.Write($"Navigation cancelled by {reason}");
    }

    internal void ProcessPointer(NavigationPointerInput input)
    {
        if (!_active) return;
        if (input.FromWinTab)
        {
            if (input.Kind == NavigationPointerKind.MiddleDown)
                ExplorerDebugLog.Write($"Navigation middle-down ignored: synthesized by WinTab at {input.Point}");
            return;
        }
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

        // Every middle click leaves a line in the log, including the ones that are not followed: a click whose
        // tab stays in the background must always be explainable afterwards.
        var startedAt = Stopwatch.GetTimestamp();
        var target = WinApi.WindowFromPoint(input.Point);
        string Describe() => $"at {input.Point.X},{input.Point.Y} delay={input.DelayMs}ms hit={target}:{WinApi.GetWindowClassName(target)} " +
            $"root={WinApi.GetAncestor(target, WinApi.GA_ROOT)}:{WinApi.GetWindowClassName(WinApi.GetAncestor(target, WinApi.GA_ROOT))} " +
            $"foreground={ExplorerNavigationAccess.ForegroundFrame()}";
        void Ignore(string reason)
        {
            CancelPending($"a new middle click ({reason})");
            ExplorerDebugLog.Write($"Navigation middle-down ignored: {reason} {Describe()}");
        }
        if (HasModifiers()) { Ignore("a modifier key is held"); return; }

        // Only window handles are read here; nothing on the input thread waits for Explorer's UI thread.
        var window = WindowAt(input.Point);
        if (window == 0) { Ignore("not in an Explorer window"); return; }
        if (!ExplorerNavigationAccess.IsInsideActiveTab(target, window) &&
            !ExplorerNavigationAccess.IsPotentialNavigationTarget(target, window))
        {
            Ignore("outside Explorer's tabs");
            return;
        }
        var onTree = ExplorerNavigationAccess.IsNavigationTree(target, window);
        var tabs = ExplorerNavigationAccess.ReadTabs(window);
        var sourceTab = tabs is { Length: > 0 } ? tabs[0] : 0;
        if (tabs is null || tabs.Length > ExplorerNavigationAccess.MaxTabs || sourceTab == 0)
        {
            Ignore($"tabs cannot be followed tabs={tabs?.Length.ToString() ?? "changing"} hwnd={window}");
            return;
        }
        var click = new NavigationClick(WindowIdentity.Capture(window), WindowIdentity.Capture(sourceTab),
            WindowIdentity.Capture(target), tabs, input.Point, onTree);
        _ = Follow(click, () => $"{Describe()} hookMs={Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F1}");
    }

    /// <summary>Follows a middle click whose window and tabs have been read, until its tab is in front or the click ends.</summary>
    internal Task<NavigationActivationResult> Follow(NavigationClick click, Func<string>? describe = null)
    {
        var lease = _clicks.Begin(click.Window.Handle, Environment.TickCount64, out var replaced);
        if (lease == null)
            return Task.FromResult(NavigationActivationResult.Cancelled);
        ExplorerDebugLog.Write($"Navigation click captured hwnd={click.Window.Handle} tab={click.SourceTab.Handle} target={click.Target.Handle} " +
            $"tree={click.OnNavigationTree} tabs={click.TabsBefore.Length} replacedEarlier={replaced} {describe?.Invoke()}");
        return Task.Run(() => ActivateAsync(click, lease));
    }

    private async Task<NavigationActivationResult> ActivateAsync(NavigationClick click, NavigationClickLease lease)
    {
        var result = NavigationActivationResult.Cancelled;
        nint loggedNewTab = 0;
        NavigationSelectOutcome? loggedOutcome = null;
        var startedAt = Stopwatch.GetTimestamp();
        string Elapsed() => $"{Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F0}ms";
        bool Current()
        {
            var current = _active && _clicks.IsCurrent(lease, Environment.TickCount64) && click.IsCurrent();
            if (!current)
                ExplorerDebugLog.Write($"Navigation request retired active={_active} lease={_clicks.IsCurrent(lease, Environment.TickCount64)} window={click.Window.IsCurrent} source={click.SourceTab.IsCurrent} sourceParent={WinApi.GetParent(click.SourceTab.Handle)} foreground={WinApi.GetForegroundWindow()}");
            return current;
        }
        try
        {
            // The only evidence of what a click did is the tab Explorer does or does not open, wherever the click
            // was: navigation pane, file list, Home page or address bar. Nothing is asked of Explorer first: an
            // accessibility query waits for Explorer's busy UI thread, which on the first click after Explorer
            // starts can take longer than the click may wait. A click that opens nothing is replaced by the next.
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
                        if (loggedNewTab != newTab)
                        {
                            ExplorerDebugLog.Write($"Navigation new tab={newTab} seen at {Elapsed()}");
                            loggedNewTab = newTab;
                        }
                        var outcome = ExplorerNavigationAccess.SelectNewTab(click, newTab, Current);
                        if (loggedOutcome != outcome)
                        {
                            ExplorerDebugLog.Write($"Navigation select outcome={outcome} at {Elapsed()}");
                            loggedOutcome = outcome;
                        }
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
            ExplorerDebugLog.Write($"Navigation activation result={result} hwnd={click.Window.Handle} at {Elapsed()}");
            _clicks.Complete(lease);
        }
        if (result == NavigationActivationResult.Activated)
            StatusChanged?.Invoke("Activated Explorer tab opened by middle-click.");
        return result;
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
