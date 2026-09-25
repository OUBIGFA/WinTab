using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using H.Hooks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Scrolling over Explorer's tab row selects the adjacent tab without wrapping. UI Automation and tab
/// commands run off the input hook; a cold hit test retains the first gesture without swallowing scrolling
/// whose target is not yet known. File-view scrolling is always left to Explorer.
/// </summary>
public sealed class ExplorerTabWheelSwitchHook : IHook
{
    private const int SelectTabCommand = 0xA221;
    private const int CommandTimeoutMs = 500;
    private const int SelectionWaitMs = 300;
    private readonly TabStripHitTester _tabStrip;
    private readonly Func<WheelSwitchSensitivity> _sensitivity;
    private readonly LowLevelMouseHook _lowLevelMouseHook;
    private readonly ExplorerTabWheelSwitchController _controller;

    public ExplorerTabWheelSwitchHook(ExplorerWatcher explorerWatcher, Func<WheelSwitchSensitivity>? sensitivity = null)
    {
        _tabStrip = explorerWatcher.TabStrip;
        _sensitivity = sensitivity ?? (() => WheelSwitchSensitivity.Medium);
        _controller = new ExplorerTabWheelSwitchController(_tabStrip.GetTabRowAsync, SwitchAsync);
        _lowLevelMouseHook = new LowLevelMouseHook { Handling = true };
        _lowLevelMouseHook.Wheel += OnWheel;
    }

    public event Action<string>? StatusChanged;
    public bool IsHookActive => _lowLevelMouseHook.IsStarted;

    public void StartHook()
    {
        _controller.Start();
        _lowLevelMouseHook.Start();
    }

    public void StopHook()
    {
        _controller.Stop();
        _lowLevelMouseHook.Stop();
    }

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
        if (e.Delta == 0) return;
        var window = GetExplorerWindowAt(e.Position);
        if (window == 0) return;
        var knownRow = _tabStrip.TryGetTabRow(window, out var tabRow);
        if (knownRow && !tabRow.Contains(e.Position)) return;

        var identity = WindowIdentity.Capture(window);
        if (!identity.IsCurrent || !WinApi.GetWindowRect(window, out var initialRect)) return;
        var point = e.Position;
        bool IsCurrent() => identity.IsCurrent && WinApi.GetWindowRect(window, out var rect) &&
            rect.Equals(initialRect) && GetExplorerWindowAt(point) == window;
        // Unknown bounds do not consume a file-view scroll. The queued input still gets its hit test and
        // tab switch once the first bounds arrive; the user does not have to turn the wheel a second time.
        if (knownRow) e.IsHandled = true;
        _ = ReportSwitchAsync(window, _controller.QueueAsync(window, point, e.Delta,
            Environment.TickCount64, _sensitivity(), IsCurrent));
    }

    private async Task ReportSwitchAsync(nint window, Task<bool> work)
    {
        try
        {
            if (await work)
                StatusChanged?.Invoke("Switched Explorer tab via mouse wheel.");
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"Wheel tab switch failed hwnd={window} error={exception.GetType().Name}");
        }
    }

    private static async Task<bool> SwitchAsync(nint window, int steps, Func<bool> isCurrent)
    {
        if (!isCurrent()) return false;
        var activeTab = ExplorerNavigationAccess.ActiveTab(window);
        if (activeTab == 0 || !ExplorerTabAutomation.TryReadSelection(window, out var currentIndex, out var tabCount) ||
            !isCurrent() || ExplorerNavigationAccess.ActiveTab(window) != activeTab ||
            ExplorerWindowDiscovery.GetAllExplorerTabs(window).Count() != tabCount ||
            ResolveTargetIndex(currentIndex, tabCount, steps) is not { } target)
            return false;

        // A timeout does not retract the command: observe Explorer's result before reporting failure.
        WinApi.TrySendMessage(window, WinApi.WM_COMMAND, SelectTabCommand, target + 1, CommandTimeoutMs);
        var deadline = Environment.TickCount64 + SelectionWaitMs;
        do
        {
            if (!isCurrent()) return false;
            var nowActive = ExplorerNavigationAccess.ActiveTab(window);
            if (nowActive != 0 && nowActive != activeTab)
                return ExplorerTabAutomation.TryReadSelection(window, out var selected, out _) && selected == target;
            await Task.Delay(5);
        }
        while (Environment.TickCount64 < deadline);

        ExplorerDebugLog.Write($"Wheel tab switch not observed hwnd={window} index={target}");
        return false;
    }

    private static nint GetExplorerWindowAt(Point point)
    {
        var hit = WinApi.WindowFromPoint(point);
        if (hit == 0) return 0;
        var root = WinApi.GetAncestor(hit, WinApi.GA_ROOT);
        return ExplorerWindowDiscovery.IsFileExplorerWindow(root) ? root : 0;
    }

    public void Dispose()
    {
        StopHook();
        _controller.Dispose();
        _lowLevelMouseHook.Dispose();
    }
}
