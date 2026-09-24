using System;
using System.Drawing;
using System.Linq;
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
    private const int SelectTabCommand = 0xA221;
    private const int CommandTimeoutMs = 500;
    private const int SelectionWaitMs = 300;
    private const int KnownSelectionTrustMs = 1_500;

    private static readonly Condition TabItemCondition =
        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);

    private readonly TabStripHitTester _tabStrip;
    private readonly Func<WheelSwitchSensitivity> _sensitivity;
    private readonly LowLevelMouseHook _lowLevelMouseHook;
    private readonly WheelSwitchThrottle _throttle = new();
    private readonly object _gate = new();
    private nint _throttleWindow;
    private nint _pendingWindow;
    private int _pendingSteps;
    private bool _workerRunning;

    // The last known selection, trusted only briefly and while the same tab is still active with the same
    // tab count; past that, a dragged tab could have moved, so the order is read again.
    private nint _knownWindow;
    private nint _knownTab;
    private int _knownIndex;
    private int _knownCount;
    private long _knownAt;

    public ExplorerTabWheelSwitchHook(ExplorerWatcher explorerWatcher, Func<WheelSwitchSensitivity>? sensitivity = null)
    {
        _tabStrip = explorerWatcher.TabStrip;
        _sensitivity = sensitivity ?? (() => WheelSwitchSensitivity.Medium);
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
        QueueStep(window, e.Delta);
    }

    private void QueueStep(nint window, int delta)
    {
        lock (_gate)
        {
            if (_throttleWindow != window)
            {
                _throttleWindow = window;
                _throttle.Reset();
            }

            var step = _throttle.Accept(delta, Environment.TickCount64, _sensitivity());
            if (step == 0)
                return;

            if (_pendingWindow != window)
                _pendingSteps = 0;
            _pendingWindow = window;
            _pendingSteps += step;
            if (_workerRunning)
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
                _knownWindow = 0;
                ExplorerDebugLog.Write($"Wheel tab switch failed hwnd={window} error={exception.GetType().Name}");
            }
        }
    }

    private async Task<bool> SwitchAsync(nint window, int steps)
    {
        if (!ExplorerWindowDiscovery.IsFileExplorerWindow(window) ||
            !TryGetSelection(window, out var activeTab, out var currentIndex, out var tabCount) ||
            ResolveTargetIndex(currentIndex, tabCount, steps) is not { } target)
            return false;

        if (!WinApi.TrySendMessage(window, WinApi.WM_COMMAND, SelectTabCommand, target + 1, CommandTimeoutMs))
        {
            _knownWindow = 0;
            return false;
        }

        // Watching the active tab handle is far cheaper than asking UI Automation again, and it tells the next
        // step exactly where the selection now is.
        var deadline = Environment.TickCount64 + SelectionWaitMs;
        while (Environment.TickCount64 < deadline)
        {
            var nowActive = GetActiveTab(window);
            if (nowActive != 0 && nowActive != activeTab)
            {
                Remember(window, nowActive, target, tabCount);
                return true;
            }

            await Task.Delay(5);
        }

        _knownWindow = 0;
        return true;
    }

    private bool TryGetSelection(nint window, out nint activeTab, out int index, out int count)
    {
        activeTab = GetActiveTab(window);
        if (window == _knownWindow && activeTab != 0 && activeTab == _knownTab &&
            Environment.TickCount64 - _knownAt <= KnownSelectionTrustMs &&
            ExplorerWindowDiscovery.GetAllExplorerTabs(window).Count() == _knownCount)
        {
            index = _knownIndex;
            count = _knownCount;
            return true;
        }

        if (!TryReadSelection(window, out index, out count))
        {
            _knownWindow = 0;
            return false;
        }

        activeTab = GetActiveTab(window);
        Remember(window, activeTab, index, count);
        return true;
    }

    private void Remember(nint window, nint tab, int index, int count)
    {
        _knownWindow = tab == 0 ? 0 : window;
        _knownTab = tab;
        _knownIndex = index;
        _knownCount = count;
        _knownAt = Environment.TickCount64;
    }

    private static nint GetActiveTab(nint window) => ExplorerNavigationAccess.ActiveTab(window);

    /// <summary>Reads the tab titles in visual order, as the tab command numbers them.</summary>
    private static bool TryReadSelection(nint window, out int selectedIndex, out int tabCount)
    {
        selectedIndex = -1;
        tabCount = 0;
        var tabItems = AutomationElement.FromHandle(window).FindAll(TreeScope.Descendants, TabItemCondition);
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
