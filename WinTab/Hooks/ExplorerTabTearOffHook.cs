using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using H.Hooks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Feeds the user's left-button presses and releases to the tab tear-off tracker. The hook only observes:
/// no input is ever swallowed, so Explorer's own drag behaviour is untouched. Events are handled on the
/// hook thread in order, which the tracker keeps cheap.
/// </summary>
internal sealed class ExplorerTabTearOffHook : IHook
{
    private readonly LowLevelMouseHook _lowLevelMouseHook;
    private readonly ExplorerTabTearOffTracker _tracker;

    public ExplorerTabTearOffHook(ExplorerTabTearOffTracker tracker)
    {
        _tracker = tracker;
        // Handling keeps delivery synchronous and in order on the hook thread; nothing is ever marked handled.
        _lowLevelMouseHook = new LowLevelMouseHook { Handling = true };
        _lowLevelMouseHook.Down += OnMouseDown;
        _lowLevelMouseHook.Up += OnMouseUp;
        _lowLevelMouseHook.ExceptionOccurred += (_, exception) =>
            ExplorerDebugLog.Write($"Tab drag hook failed: {exception.GetType().Name}:{exception.Message}");
    }

    public bool IsHookActive => _lowLevelMouseHook.IsStarted;

    public void StartHook()
    {
        if (!_lowLevelMouseHook.IsStarted)
            _lowLevelMouseHook.Start();
    }

    public void StopHook() => _lowLevelMouseHook.Stop();

    private void OnMouseDown(object? sender, MouseEventArgs e)
    {
        if (IsLeftMouse(e))
            Observe(() => _tracker.HandleLeftDown(e.Position));
    }

    private void OnMouseUp(object? sender, MouseEventArgs e)
    {
        if (IsLeftMouse(e))
            Observe(() => _tracker.HandleLeftUp(e.Position, Environment.TickCount64));
    }

    private static void Observe(Action observation)
    {
        try
        {
            observation();
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"Tab drag observation failed: {exception.GetType().Name}:{exception.Message}");
        }
    }

    private static bool IsLeftMouse(MouseEventArgs e) => e.CurrentKey is Key.MouseLeft or Key.LButton;

    public void Dispose()
    {
        StopHook();
        _lowLevelMouseHook.Dispose();
    }
}

/// <summary>The real screen: Explorer windows, their tab titles and rows, and their bounds.</summary>
internal sealed class ExplorerTabTearOffEnvironment(TabStripHitTester tabStrip, Func<bool> isEnabled) : IExplorerTabTearOffEnvironment
{
    public bool IsEnabled => isEnabled();
    public int DragWidth => WinApi.GetSystemMetrics(WinApi.SM_CXDRAG);
    public int DragHeight => WinApi.GetSystemMetrics(WinApi.SM_CYDRAG);

    public nint ResolveExplorerWindow(Point point)
    {
        var hit = WinApi.WindowFromPoint(point);
        if (hit == 0)
            return 0;
        var root = WinApi.GetAncestor(hit, WinApi.GA_ROOT);
        return ExplorerWindowDiscovery.IsFileExplorerWindow(root) ? root : 0;
    }

    public bool IsPointOnTab(Point point, nint explorerWindow) => tabStrip.IsPointOnTabStrip(point, explorerWindow);

    public Rectangle GetTabRow(nint explorerWindow) =>
        tabStrip.TryGetTabRow(explorerWindow, out var tabRow) ? tabRow : Rectangle.Empty;

    public int CountTabs(nint explorerWindow) => ExplorerWindowDiscovery.GetAllExplorerTabs(explorerWindow).Count();

    public bool IsWindowShown(nint explorerWindow) => WinApi.IsWindowVisible(explorerWindow);

    public IEnumerable<nint> ListShownExplorerWindows() => ExplorerWindowDiscovery.GetAllExplorerWindows().Where(WinApi.IsWindowVisible);

    public Rectangle GetWindowBounds(nint explorerWindow) =>
        WinApi.GetWindowRect(explorerWindow, out var rect)
            ? Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom)
            : Rectangle.Empty;

    public Func<bool> TrackIdentity(nint explorerWindow)
    {
        var identity = WindowIdentity.Capture(explorerWindow);
        return () => identity.IsCurrent;
    }
}
