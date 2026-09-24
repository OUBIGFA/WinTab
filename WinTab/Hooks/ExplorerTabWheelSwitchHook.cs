using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using H.Hooks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Scrolling the mouse wheel over an Explorer window's tab row switches its active tab: wheel up selects
/// the tab to the left, wheel down the tab to the right, stopping at either end. The wheel is swallowed only
/// over the tab row; everywhere else it reaches Explorer untouched.
/// </summary>
public sealed class ExplorerTabWheelSwitchHook : IHook
{
    private const int WheelDelta = 120;
    private const int SelectTabCommand = 0xA221;
    private const int CommandTimeoutMs = 500;

    private readonly TabStripHitTester _tabStrip;
    private readonly LowLevelMouseHook _lowLevelMouseHook;
    private readonly object _gate = new();
    private nint _pendingWindow;
    private int _pendingSteps;
    private int _wheelRemainder;
    private bool _workerRunning;

    public ExplorerTabWheelSwitchHook(ExplorerWatcher explorerWatcher)
    {
        _tabStrip = explorerWatcher.TabStrip;
        _lowLevelMouseHook = new LowLevelMouseHook { Handling = true };
        _lowLevelMouseHook.Wheel += OnWheel;
    }

    public event Action<string>? StatusChanged;
    public bool IsHookActive => _lowLevelMouseHook.IsStarted;

    public void StartHook() => _lowLevelMouseHook.Start();
    public void StopHook() => _lowLevelMouseHook.Stop();

    /// <summary>The tab index a wheel movement lands on, or null when it would not change the selection.</summary>
    internal static int? ResolveTargetIndex(int currentIndex, int tabCount, int steps)
    {
        if (tabCount < 2 || currentIndex < 0 || currentIndex >= tabCount || steps == 0)
            return null;

        var target = Math.Clamp(currentIndex + steps, 0, tabCount - 1);
        return target == currentIndex ? null : target;
    }

    private void OnWheel(object? sender, MouseEventArgs e)
    {
        if (e.Delta == 0)
            return;

        var window = GetExplorerWindowAt(e.Position);
        if (window == 0)
            return;

        if (!_tabStrip.TryGetTabRow(window, out var tabRow))
        {
            _tabStrip.ScheduleRefresh(window);
            return;
        }

        if (!tabRow.Contains(e.Position))
            return;

        e.IsHandled = true;
        QueueSteps(window, e.Delta);
    }

    private void QueueSteps(nint window, int delta)
    {
        lock (_gate)
        {
            if (_pendingWindow != window)
            {
                _pendingWindow = window;
                _pendingSteps = 0;
                _wheelRemainder = 0;
            }

            // High-resolution wheels report fractions of a notch; only whole notches move the selection.
            // Wheel up (positive delta) goes to the previous tab, matching the direction the list scrolls.
            _wheelRemainder += delta;
            var notches = _wheelRemainder / WheelDelta;
            _wheelRemainder -= notches * WheelDelta;
            _pendingSteps -= notches;
            if (_pendingSteps == 0 || _workerRunning)
                return;

            _workerRunning = true;
        }

        // UI Automation and the tab command both wait on Explorer; keep them off the hook thread.
        Task.Run(DrainAsync);
    }

    private async Task DrainAsync()
    {
        while (true)
        {
            nint window;
            int steps;
            lock (_gate)
            {
                if (_pendingSteps == 0)
                {
                    _workerRunning = false;
                    return;
                }

                window = _pendingWindow;
                steps = _pendingSteps;
                _pendingSteps = 0;
            }

            try
            {
                if (await SwitchAsync(window, steps))
                    StatusChanged?.Invoke("Switched Explorer tab via mouse wheel.");
            }
            catch (Exception exception)
            {
                ExplorerDebugLog.Write($"Wheel tab switch failed hwnd={window} error={exception.GetType().Name}");
            }
        }
    }

    private static async Task<bool> SwitchAsync(nint window, int steps)
    {
        if (!ExplorerWindowDiscovery.IsFileExplorerWindow(window) ||
            !TryGetSelectedTab(window, out var currentIndex, out var tabCount) ||
            ResolveTargetIndex(currentIndex, tabCount, steps) is not { } target)
            return false;

        if (!WinApi.TrySendMessage(window, WinApi.WM_COMMAND, SelectTabCommand, target + 1, CommandTimeoutMs))
            return false;

        // Give Explorer a moment to move the selection before the next queued notch reads it again.
        for (var attempt = 0; attempt < 20; attempt++)
        {
            if (TryGetSelectedTab(window, out var selected, out _) && selected == target)
                return true;
            await Task.Delay(10);
        }

        return true;
    }

    /// <summary>Reads the tab titles in visual order, as the tab command numbers them.</summary>
    private static bool TryGetSelectedTab(nint window, out int selectedIndex, out int tabCount)
    {
        selectedIndex = -1;
        tabCount = 0;
        var tabItems = AutomationElement.FromHandle(window).FindAll(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
        foreach (AutomationElement tab in tabItems)
        {
            if (tab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var pattern) &&
                ((SelectionItemPattern)pattern).Current.IsSelected)
                selectedIndex = tabCount;
            tabCount++;
        }

        return selectedIndex >= 0;
    }

    private static nint GetExplorerWindowAt(Point point)
    {
        var hit = WinApi.WindowFromPoint(point);
        if (hit == 0)
            return 0;

        var root = WinApi.GetAncestor(hit, WinApi.GA_ROOT);
        return ExplorerWindowDiscovery.IsFileExplorerWindow(root) ? root : 0;
    }

    public void Dispose()
    {
        StopHook();
        _lowLevelMouseHook.Dispose();
    }
}
