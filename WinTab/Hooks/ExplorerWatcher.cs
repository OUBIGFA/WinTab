using Shell32;
using SHDocVw;
using System;
using System.Linq;
using System.IO;
using System.Windows;
using System.Threading;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using WinTab.Helpers;
using WinTab.Interop;
using WinTab.Managers;
using WinTab.Models;
using WinTab.WinAPI;
using WinTab.UI.Views;

namespace WinTab.Hooks;

using WindowEntry = DualKeyEntry<InternetExplorer, nint?, WindowInfo>;
using DrawingPoint = System.Drawing.Point;
using DrawingRectangle = System.Drawing.Rectangle;

public class ExplorerWatcher : IHook
{
    private const int AutomationRootTtlMs = 5_000;
    private const int NavigationCompleteWaitMs = 1_200;
    private const int NavigationVerificationWaitMs = 1_200;
    private const int StripBoundsRectMatchSlop = 2;
    private const int StartupLocationCacheLimit = 512;
    private static readonly string? DebugLogPath = Environment.GetEnvironmentVariable("WINTAB_DEBUG_LOG");
    private static bool _instanceRunning;
    private static Guid _shellBrowserGuid = typeof(IShellBrowser).GUID;

    private ShellWindows _shellWindows = null!;
    private ShellPathComparer _shellPathComparer = null!;
    private StaTaskScheduler _staTaskScheduler = null!;
    private nint _mainWindowHandle;
    private readonly ConcurrentDictionary<nint, byte> _processedHWnds = new();
    private readonly ConcurrentDictionary<nint, int> _hookedTopLevelUseCounts = new();
    private readonly ConcurrentDictionary<nint, TabStripBounds> _stripBoundsCache = new();
    private readonly ConcurrentDictionary<nint, byte> _stripBoundsRefreshInflight = new();
    private readonly ConcurrentDictionary<nint, (AutomationElement Element, long ExpiresAt)> _automationRootCache = new();
    private readonly ConcurrentDictionary<string, bool> _startupLocationCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly DualKeyDictionary<InternetExplorer, nint?, WindowInfo> _windowEntryDict = [];
    private readonly List<WindowRecord> _closedWindows = new();
    private readonly object _windowEntryDictLock = new(), _closedWindowsLock = new(), _processLock = new();
    private readonly SemaphoreSlim _toOpenWindowsLock = new(1);
    private readonly SemaphoreSlim _shellWindowRegistrationLock = new(1);
    private readonly ExplorerLaunchLocationResolver _locationResolver = new();
    private readonly ProcessWatcher _processWatcher;
    private int _mainExplorerProcessId;
    private Timer? _explorerCheckTimer;

    private WinEventHookThread? _winEventHookThread;
    private WinEventDelegate? _eventObjectShowHookCallback;
    private DShellWindowsEvents_WindowRegisteredEventHandler? _windowRegisteredHandler;

    private string _defaultLocation = null!;
    private bool _reuseTabs = true;
    private bool _isForcingTabs;
    private int _shellWindowRegistrationScheduled;
    private int _mergeSourceConcealPulseRunning;
    private long _mergeSourceConcealPulseUntilTicks;
    private readonly ConcurrentDictionary<nint, byte> _mergeSourceHWnds = new();
    private sealed record TabStripBounds(DrawingRectangle StripRect, DrawingRectangle[] TabRects, RECT WindowRect, long RefreshedAt);
    public bool IsHookActive => _isForcingTabs;
    public bool IsShellReady => _mainExplorerProcessId != 0 && _shellWindows != null;
    public event Action? OnShellInitialized;

    private static void DebugLog(string message)
    {
        if (string.IsNullOrWhiteSpace(DebugLogPath))
            return;

        try
        {
            File.AppendAllText(DebugLogPath, $"{DateTime.Now:HH:mm:ss.fff} {message}{Environment.NewLine}");
        }
        catch
        {
            //
        }
    }

    public ExplorerWatcher()
    {
        if (_instanceRunning)
            throw new InvalidOperationException("Only one instance of ExplorerWatcher is allowed at a time.");
        _instanceRunning = true;

        _processWatcher = new ProcessWatcher("explorer");
        _processWatcher.ProcessTerminated += OnExplorerProcessTerminated;
        StartExplorerProcessCheck();
    }

    public void StartHook()
    {
        if (_isForcingTabs) return;
        RecoverHiddenExplorerWindows("start-hook");
        _isForcingTabs = true;
        StartMergeSourceConcealPulse(500);
        DebugLog("StartHook");
    }

    public void StopHook()
    {
        if (!_isForcingTabs) return;
        _isForcingTabs = false;
        StopMergeSourceConcealPulse();
        RecoverHiddenExplorerWindows("stop-hook");
    }
    public void SetReuseTabs(bool reuseTabs) => _reuseTabs = reuseTabs;

    public bool IsPointOnTabStrip(DrawingPoint screenPoint, nint explorerWindow)
    {
        if (!Helper.IsFileExplorerWindow(explorerWindow))
            return false;

        if (!WinApi.GetWindowRect(explorerWindow, out var winRect))
            return false;

        if (screenPoint.X < winRect.Left || screenPoint.X >= winRect.Right || screenPoint.Y < winRect.Top)
            return false;

        if (_stripBoundsCache.TryGetValue(explorerWindow, out var bounds) &&
            RectsApproxEqual(bounds.WindowRect, winRect))
        {
            if (Environment.TickCount64 - bounds.RefreshedAt > 1_500)
                ScheduleStripBoundsRefresh(explorerWindow);

            return bounds.TabRects.Any(tabRect => tabRect.Contains(screenPoint.X, screenPoint.Y));
        }

        if (bounds != null)
            _stripBoundsCache.TryRemove(explorerWindow, out _);

        ScheduleStripBoundsRefresh(explorerWindow);
        return false;
    }

    public void RefreshTabStripBounds(nint explorerWindow)
    {
        InvalidateStripBounds(explorerWindow);
        if (Helper.IsFileExplorerWindow(explorerWindow))
            ScheduleStripBoundsRefresh(explorerWindow);
    }

    private static bool RectsApproxEqual(RECT a, RECT b)
    {
        return Math.Abs(a.Left - b.Left) <= StripBoundsRectMatchSlop &&
               Math.Abs(a.Top - b.Top) <= StripBoundsRectMatchSlop &&
               Math.Abs(a.Right - b.Right) <= StripBoundsRectMatchSlop &&
               Math.Abs(a.Bottom - b.Bottom) <= StripBoundsRectMatchSlop;
    }

    private void ScheduleStripBoundsRefresh(nint explorerWindow)
    {
        if (!_stripBoundsRefreshInflight.TryAdd(explorerWindow, 0))
            return;

        Task.Run(() =>
        {
            try
            {
                ComputeStripBounds(explorerWindow);
            }
            catch
            {
                //
            }
            finally
            {
                _stripBoundsRefreshInflight.TryRemove(explorerWindow, out _);
            }
        });
    }

    private void ComputeStripBounds(nint explorerWindow)
    {
        if (!Helper.IsFileExplorerWindow(explorerWindow))
            return;

        var root = GetAutomationRoot(explorerWindow);
        if (root == null)
            return;

        var tabItems = root.FindAll(
            TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
        if (tabItems == null || tabItems.Count == 0)
            return;

        var tabRects = new List<DrawingRectangle>(tabItems.Count);
        AutomationElement? firstTabItem = null;
        foreach (AutomationElement tab in tabItems)
        {
            var rect = tab.Current.BoundingRectangle;
            if (rect.IsEmpty || rect.Width < 24 || rect.Height < 12)
                continue;

            firstTabItem ??= tab;
            tabRects.Add(new DrawingRectangle((int)rect.X, (int)rect.Y, (int)rect.Width, (int)rect.Height));
        }

        if (firstTabItem == null || tabRects.Count == 0)
            return;

        var walker = TreeWalker.ControlViewWalker;
        var element = firstTabItem;
        Rect? stripRect = null;
        while (element != null)
        {
            if (element.Current.ControlType == ControlType.Tab)
            {
                var rect = element.Current.BoundingRectangle;
                if (!rect.IsEmpty)
                {
                    stripRect = rect;
                    break;
                }
            }

            element = walker.GetParent(element);
        }

        if (!stripRect.HasValue)
            stripRect = firstTabItem.Current.BoundingRectangle;

        if (!stripRect.HasValue || stripRect.Value.IsEmpty)
            return;

        if (!WinApi.GetWindowRect(explorerWindow, out var winSnapshot))
            return;

        var strip = stripRect.Value;
        _stripBoundsCache[explorerWindow] = new TabStripBounds(
            new DrawingRectangle((int)strip.X, (int)strip.Y, (int)strip.Width, (int)strip.Height),
            tabRects.ToArray(),
            winSnapshot,
            Environment.TickCount64);
    }

    private AutomationElement? GetAutomationRoot(nint explorerWindow)
    {
        var now = Environment.TickCount64;
        if (_automationRootCache.TryGetValue(explorerWindow, out var entry) && entry.ExpiresAt > now)
            return entry.Element;

        try
        {
            var root = AutomationElement.FromHandle(explorerWindow);
            if (root != null)
                _automationRootCache[explorerWindow] = (root, now + AutomationRootTtlMs);

            return root;
        }
        catch
        {
            return null;
        }
    }

    private void InvalidateStripBounds(nint hWnd)
    {
        _stripBoundsCache.TryRemove(hWnd, out _);
        _stripBoundsRefreshInflight.TryRemove(hWnd, out _);
    }

    private void InvalidateAutomationRoot(nint hWnd)
    {
        _automationRootCache.TryRemove(hWnd, out _);
    }

    public void ClearClosedWindows()
    {
        lock (_closedWindowsLock)
            _closedWindows.Clear();
    }

    public IReadOnlyCollection<WindowRecord> GetWindows()
    {
        var result = new List<WindowRecord>();
        
        // Add open windows
        lock (_windowEntryDictLock)
            result.AddRange(
                _windowEntryDict.Keys.Select(ie => new WindowRecord(GetLocation(ie), new IntPtr(ie.HWND), GetSelectedItems(ie), ie.LocationName)));
        
        // Add closed windows in reverse order (last closed on top)
        lock (_closedWindowsLock)
            result.AddRange(_closedWindows.AsEnumerable().Reverse());
        
        return result.GroupBy(w => w.Location).Select(g => g.First()).ToList();
    }

    public async Task SwitchTo(string location, nint windowHandle = 0, string[]? selectedItems = null, bool asTab = true, bool duplicate = false)
    {
        var windowToOpen = new WindowRecord(location, windowHandle, selectedItems);
        if (!asTab)
        {
            await OpenNewWindowWithSelection(windowToOpen);
            return;
        }

        await OpenTabNavigateWithSelection(windowToOpen, windowHandle, duplicate, true);
    }
    
    public nint SearchForTab(string targetPath)
    {
        return TrySearchForTab(targetPath, excludedTopLevelWindow: 0, out var tabHandle, out _) ? tabHandle : 0;
    }
    private bool TrySearchForTab(string targetPath, nint excludedTopLevelWindow, out nint tabHandle, out InternetExplorer? foundWindow)
    {
        nint targetPidl = 0;
        tabHandle = 0;
        foundWindow = null;
        try
        {
            targetPidl = _shellPathComparer.GetPidlFromPath(targetPath);
            if (targetPidl == 0) return false;

            foreach (var (window, windowInfo, tab) in _windowEntryDict)
            {
                if (!windowInfo.EventsHooked ||
                    !tab.HasValue ||
                    tab.Value == 0)
                {
                    continue;
                }

                nint topLevelWindow;
                try
                {
                    topLevelWindow = new IntPtr(window.HWND);
                }
                catch
                {
                    continue;
                }

                if (excludedTopLevelWindow != 0 && topLevelWindow == excludedTopLevelWindow)
                    continue;

                var comparePath = windowInfo.Location ?? GetLocation(window);

                if (_shellPathComparer.IsEquivalent(targetPath, comparePath, targetPidl))
                {
                    foundWindow = window;
                    tabHandle = tab.Value;
                    return true;
                }
            }

            return false;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (targetPidl != 0)
                Marshal.FreeCoTaskMem(targetPidl);
        }
    }
    public async Task SelectTabByHandle(nint windowHandle, nint tabHandle)
    {
        await TrySelectTabByHandleDirectAsync(windowHandle, tabHandle, timeoutMs: 500);
    }
    public void SelectLastTab(nint windowHandle)
    {
        var count = Helper.GetAllExplorerTabs(windowHandle).Count();
        SelectTabByIndex(windowHandle, count - 1);
    }
    public void SelectTabByIndex(nint windowHandle, int index)
    {
        // Send 0xA221 magic command (CTRL + 1...n)
        WinApi.SendMessage(windowHandle, WinApi.WM_COMMAND, 0xA221, index + 1);
    }
    private async Task<bool> TrySelectTabByHandleDirectAsync(nint windowHandle, nint tabHandle, int timeoutMs)
    {
        if (windowHandle == 0 || tabHandle == 0)
            return false;

        if (GetActiveTabHandle(windowHandle) == tabHandle)
            return true;

        return await SelectTabByUniqueNameVerified(windowHandle, tabHandle, timeoutMs);
    }
    private async Task<bool> SelectTabByUniqueNameVerified(nint windowHandle, nint tabHandle, int timeoutMs, InternetExplorer? knownWindow = null)
    {
        if (GetActiveTabHandle(windowHandle) == tabHandle)
            return true;

        var window = knownWindow ?? GetWindowByTabHandle(tabHandle) ?? await FindShellWindowByTabHandle(tabHandle, windowHandle);
        if (window == null)
            return false;

        string tabName;
        try
        {
            tabName = window.LocationName;
        }
        catch
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(tabName))
            return false;

        if (TrySelectSingleTabByAutomationName(windowHandle, tabName))
        {
            var activeTab = await Helper.DoUntilConditionAsync(
                () => GetActiveTabHandle(windowHandle),
                h => h == tabHandle,
                timeoutMs,
                10);

            if (activeTab == tabHandle)
                return true;
        }

        InvalidateAutomationRoot(windowHandle);
        if (TrySelectSingleTabByAutomationName(windowHandle, tabName))
        {
            var activeTab = await Helper.DoUntilConditionAsync(
                () => GetActiveTabHandle(windowHandle),
                h => h == tabHandle,
                timeoutMs,
                10);

            if (activeTab == tabHandle)
                return true;
        }

        return await TrySelectTabByCyclingAsync(windowHandle, tabHandle, timeoutMs);
    }
    private async Task<bool> TrySelectTabByCyclingAsync(nint windowHandle, nint tabHandle, int timeoutMs)
    {
        var tabs = Helper.GetAllExplorerTabs(windowHandle).ToArray();
        if (tabs.Length == 0)
            return false;

        var activeTab = GetActiveTabHandle(windowHandle);
        var perStepTimeout = Math.Max(50, timeoutMs / Math.Max(1, tabs.Length));
        for (var i = 0; i < tabs.Length; i++)
        {
            if (activeTab == tabHandle)
                return true;

            SelectTabByIndex(windowHandle, i);
            var previousTab = activeTab;
            activeTab = await Helper.DoUntilConditionAsync(
                () => GetActiveTabHandle(windowHandle),
                h => h == tabHandle || h != previousTab,
                perStepTimeout,
                10);
        }

        return GetActiveTabHandle(windowHandle) == tabHandle;
    }
    private bool TrySelectSingleTabByAutomationName(nint windowHandle, string tabName)
    {
        try
        {
            var root = GetAutomationRoot(windowHandle);
            if (root == null)
                return false;

            var tabItems = root.FindAll(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
            AutomationElement? matchingTab = null;

            foreach (AutomationElement element in tabItems)
            {
                var rect = element.Current.BoundingRectangle;
                if (rect.IsEmpty || rect.Width < 24 || rect.Height < 12)
                    continue;

                if (!IsMatchingTabName(element.Current.Name, tabName))
                    continue;

                if (matchingTab != null)
                    return false;

                matchingTab = element;
            }

            return matchingTab != null && TrySelectAutomationElement(matchingTab);
        }
        catch
        {
            return false;
        }
    }
    private static bool IsMatchingTabName(string automationName, string tabName)
    {
        if (string.IsNullOrWhiteSpace(automationName) || string.IsNullOrWhiteSpace(tabName))
            return false;

        if (StringComparer.OrdinalIgnoreCase.Equals(automationName.Trim(), tabName.Trim()))
            return true;

        return automationName.Contains(tabName, StringComparison.OrdinalIgnoreCase);
    }
    private static bool TrySelectAutomationElement(AutomationElement element)
    {
        if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectionPattern) &&
            selectionPattern is SelectionItemPattern selectionItemPattern)
        {
            selectionItemPattern.Select();
            return true;
        }

        if (element.TryGetCurrentPattern(InvokePattern.Pattern, out var invokePattern) &&
            invokePattern is InvokePattern invoke)
        {
            invoke.Invoke();
            return true;
        }

        return false;
    }
    public async Task RequestToOpenNewTab(nint windowHandle, bool bringToFront = false, bool lockToOpenWindows = true)
    {
        if (bringToFront && windowHandle == 0)
            windowHandle = GetMainWindowHWnd(0);

        if (windowHandle == 0)
        {
            await OpenNewWindowWithSelection(new WindowRecord(string.Empty), lockToOpenWindows);
            return;
        }

        var tabHandle = WinApi.FindWindowEx(windowHandle, 0, "ShellTabWindowClass", null);
        if (tabHandle == 0) return;

        // Send 0xA21B magic command (CTRL + T)
        WinApi.PostMessage(tabHandle, WinApi.WM_COMMAND, 0xA21B, 0);

        if (bringToFront)
            WinApi.RestoreWindowToForeground(windowHandle);
    }
    public async Task Open(string? location, bool asTab, nint windowHandle, int delay = 0)
    {
        if (delay > 0)
            await Task.Delay(delay);

        var normalizedPath = Helper.NormalizeLocation(location ?? string.Empty);
        
        if (normalizedPath.StartsWith("http", StringComparison.OrdinalIgnoreCase) ||
            normalizedPath.StartsWith("www.", StringComparison.OrdinalIgnoreCase) ||
            System.IO.File.Exists(normalizedPath))
        {
            try
            {
                Helper.BypassWinForegroundRestrictions();
                Process.Start(new ProcessStartInfo(normalizedPath) { UseShellExecute = true });
                return;
            }
            catch
            {
                //
            }
        }
        
        if (!asTab)
        {
            await OpenNewWindowWithSelection(new WindowRecord(normalizedPath));
            return;
        }

        if (string.IsNullOrWhiteSpace(normalizedPath) && !_reuseTabs)
        {
            await RequestToOpenNewTab(windowHandle, bringToFront: true);
            return;
        }

        if (_windowEntryDict.Count > 0)
        {
            OpenNewTab(windowHandle, normalizedPath);
            return;
        }

        await OpenNewWindowWithSelection(new WindowRecord(normalizedPath));
    }
    public void OpenNewTab(nint windowHandle, string location)
    {
        _ = OpenTabNavigateWithSelection(new WindowRecord(location, windowHandle), windowHandle);
    }
    public async Task DuplicateActiveTab(nint windowHandle, bool asTab)
    {
        var activeTabHandle = GetActiveTabHandle(windowHandle);
        if (activeTabHandle == 0) return;

        var window = GetWindowByTabHandle(activeTabHandle);
        if (window == null) return;

        var location = _windowEntryDict[window].Value.Location ?? GetLocation(window);
        var selectedItems = GetSelectedItems(window);
        var windowRecord = new WindowRecord(location, windowHandle, selectedItems);

        if (!asTab)
        {
            await OpenNewWindowWithSelection(windowRecord);
            return;
        }

        await OpenTabNavigateWithSelection(windowRecord, windowHandle, isDuplicate: true);
    }
    public async Task ReopenClosedTab(bool asTab, nint windowHandle = 0)
    {
        WindowRecord? closedWindow;
        lock (_closedWindowsLock)
        {
            closedWindow = _closedWindows.LastOrDefault(w => w.Location != _defaultLocation);
            if (closedWindow == null) return;
            _closedWindows.Remove(closedWindow);
        }

        if (!asTab)
        {
            closedWindow.CreatedAt = Environment.TickCount;
            await OpenNewWindowWithSelection(closedWindow);
            return;
        }

        await OpenTabNavigateWithSelection(closedWindow, windowHandle);
    }
    public async Task DetachCurrentTab(nint windowHandle)
    {
        if (Helper.GetAllExplorerTabs(windowHandle).Take(2).Count() < 2)
            return;

        var activeTabHandle = GetActiveTabHandle(windowHandle);
        if (activeTabHandle == 0) return;

        var window = GetWindowByTabHandle(activeTabHandle);
        if (window == null) return;

        var location = _windowEntryDict[window].Value.Location ?? GetLocation(window);
        var selectedItems = GetSelectedItems(window);
        var windowRecord = new WindowRecord(location, windowHandle, selectedItems);

        // Send 0xA021 magic command (CTRL + W)
        WinApi.SendMessage(activeTabHandle, WinApi.WM_COMMAND, 0xA021, 1);

        await OpenNewWindowWithSelection(windowRecord);
    }
    public void SetTargetWindow(nint windowHandle)
    {
        if (Helper.IsFileExplorerWindow(windowHandle))
            _mainWindowHandle = windowHandle;
    }
    public void NavigateBackForward(nint windowHandle, bool isForward)
    {
        var activeTabHandle = GetActiveTabHandle(windowHandle);
        if (activeTabHandle == 0) return;

        var window = GetWindowByTabHandle(activeTabHandle);
        try
        {
            if (isForward) window?.GoForward();
            else window?.GoBack();
        }
        catch
        {
            // Will throw if there is no further history
        }
    }

    private void PreventWindowHiding(nint hWnd)
    {
        if (_processedHWnds.TryAdd(hWnd, 0))
        {
            // Schedule removal after a short delay
            _ = Task.Delay(7_000).ContinueWith(t => _processedHWnds.TryRemove(hWnd, out _), TaskScheduler.Default);
        }
    }
    private void OnWindowShown(nint hWinEventHook, uint eventType, nint hWnd, int idObject, int idChild, uint dwEventThread, uint dWmsEventTime)
    {
        if (!_isForcingTabs || hWnd == 0) return;

        TryHideIncomingExplorerWindow(hWnd);
        StartMergeSourceConcealPulse();
        ScheduleShellWindowRegistration(1);
    }
    private bool TryHideIncomingExplorerWindow(nint hWnd)
    {
        hWnd = GetExplorerTopLevelWindow(hWnd);
        if ((!_isForcingTabs && !_reuseTabs) || hWnd == 0) return false;
        if (_processedHWnds.ContainsKey(hWnd)) return false;
        if (Helper.IsCtrlShiftDown()) return false;
        if (_hookedTopLevelUseCounts.ContainsKey(hWnd)) return false;
        if (_mainWindowHandle != 0 && hWnd == _mainWindowHandle) return false;
        if (Helper.GetAllExplorerTabs(hWnd).Take(2).Count() > 1) return false;

        HideMergeSourceWindow(hWnd);
        return true;
    }
    private static nint GetExplorerTopLevelWindow(nint hWnd)
    {
        if (hWnd == 0)
            return 0;

        if (WinApi.IsWindowHasClassName(hWnd, "CabinetWClass"))
            return hWnd;

        var root = WinApi.GetAncestor(hWnd, WinApi.GA_ROOT);
        return root != 0 && WinApi.IsWindowHasClassName(root, "CabinetWClass")
            ? root
            : 0;
    }
    private void TryHideRegisteredMergeSourceWindow(nint hWnd)
    {
        if ((!_isForcingTabs && !_reuseTabs) || hWnd == 0) return;
        if (_processedHWnds.ContainsKey(hWnd)) return;
        if (Helper.IsCtrlShiftDown()) return;
        if (_mainWindowHandle != 0 && hWnd == _mainWindowHandle) return;
        if (Helper.GetAllExplorerTabs(hWnd).Take(2).Count() > 1) return;

        var targetWindow = GetMainWindowHWnd(hWnd);
        if (targetWindow == 0 || hWnd == targetWindow) return;

        HideMergeSourceWindow(hWnd);
    }
    private void HideMergeSourceWindow(nint hWnd)
    {
        if (hWnd == 0)
            return;

        var firstHide = _mergeSourceHWnds.TryAdd(hWnd, 0);
        Helper.HideWindow(hWnd, SettingsManager.HaveThemeIssue);
        if (firstHide)
            DebugLog($"Concealed merge source hwnd={hWnd}");
    }
    private async Task RestoreMergeSourceWindowAsync(nint hWnd)
    {
        if (hWnd == 0)
            return;

        PreventWindowHiding(hWnd);
        if (!_mergeSourceHWnds.TryRemove(hWnd, out _) && !Helper.HiddenWindows.ContainsKey(hWnd))
            return;

        await RestoreHiddenExplorerWindowAsync(hWnd);
        DebugLog($"Restored merge source hwnd={hWnd}");
    }
    private void RemoveMergeSourceTracking(nint hWnd)
    {
        if (hWnd == 0)
            return;

        _mergeSourceHWnds.TryRemove(hWnd, out _);
        Helper.HiddenWindows.TryRemove(hWnd, out _);
        PreventWindowHiding(hWnd);
    }
    private void StartMergeSourceConcealPulse(int durationMs = 1_200)
    {
        if (!_isForcingTabs && !_reuseTabs)
            return;

        var untilTicks = DateTime.UtcNow.AddMilliseconds(Math.Max(1, durationMs)).Ticks;
        var currentTicks = Volatile.Read(ref _mergeSourceConcealPulseUntilTicks);
        while (untilTicks > currentTicks)
        {
            var previousTicks = Interlocked.CompareExchange(
                ref _mergeSourceConcealPulseUntilTicks,
                untilTicks,
                currentTicks);
            if (previousTicks == currentTicks)
                break;

            currentTicks = previousTicks;
        }

        if (Interlocked.Exchange(ref _mergeSourceConcealPulseRunning, 1) == 1)
            return;

        _ = Task.Run(async () =>
        {
            try
            {
                while (DateTime.UtcNow.Ticks < Volatile.Read(ref _mergeSourceConcealPulseUntilTicks))
                {
                    try
                    {
                        ConcealMergeSourceWindowsOnce();
                    }
                    catch
                    {
                        //
                    }

                    await Task.Delay(15);
                }
            }
            finally
            {
                Interlocked.Exchange(ref _mergeSourceConcealPulseRunning, 0);
                if (DateTime.UtcNow.Ticks < Volatile.Read(ref _mergeSourceConcealPulseUntilTicks))
                    StartMergeSourceConcealPulse(250);
            }
        });
    }
    private void ConcealMergeSourceWindowsOnce()
    {
        foreach (var hWnd in WinApi.FindAllWindowsEx("CabinetWClass"))
        {
            try
            {
                TryHideIncomingExplorerWindow(hWnd);
            }
            catch
            {
                //
            }
        }
    }
    private void StopMergeSourceConcealPulse()
    {
        Interlocked.Exchange(ref _mergeSourceConcealPulseUntilTicks, 0);
    }
    private bool HasHookedShellWindowForTopLevel(nint hWnd)
    {
        lock (_windowEntryDictLock)
        {
            foreach (var (window, info, _) in _windowEntryDict)
            {
                if (!info.EventsHooked)
                    continue;

                try
                {
                    if (new IntPtr(window.HWND) == hWnd)
                        return true;
                }
                catch
                {
                    //
                }
            }
        }

        return false;
    }
    private bool HasTrackedTopLevelWindow(nint hWnd)
    {
        lock (_windowEntryDictLock)
        {
            foreach (var (window, _, _) in _windowEntryDict)
            {
                try
                {
                    if (new IntPtr(window.HWND) == hWnd)
                        return true;
                }
                catch
                {
                    //
                }
            }
        }

        return false;
    }
    private bool HasOtherTrackedShellWindowForTopLevel(InternetExplorer currentWindow, nint hWnd)
    {
        lock (_windowEntryDictLock)
        {
            foreach (var (window, _, _) in _windowEntryDict)
            {
                if (ReferenceEquals(window, currentWindow))
                    continue;

                try
                {
                    if (new IntPtr(window.HWND) == hWnd)
                        return true;
                }
                catch
                {
                    //
                }
            }
        }

        return false;
    }
    private void ReleaseTopLevelTrackingIfUnused(nint hWnd)
    {
        if (hWnd == 0)
            return;

        if (HasTrackedTopLevelWindow(hWnd) || Helper.IsFileExplorerWindow(hWnd))
        {
            _ = Task.Delay(5_000).ContinueWith(_ =>
            {
                if (!HasTrackedTopLevelWindow(hWnd) && !Helper.IsFileExplorerWindow(hWnd))
                    ReleaseTopLevelTracking(hWnd);
            }, TaskScheduler.Default);
            return;
        }

        ReleaseTopLevelTracking(hWnd);
    }
    private void ReleaseTopLevelTracking(nint hWnd)
    {
        _processedHWnds.TryRemove(hWnd, out _);
    }
    private sealed class WinEventHookThread : IDisposable
    {
        private readonly WinEventDelegate _showCallback;
        private readonly ManualResetEventSlim _started = new();
        private readonly ManualResetEventSlim _stopped = new();
        private Thread? _thread;
        private uint _threadId;
        private nint _foregroundHookId;
        private nint _showHookId;
        private bool _disposed;

        public WinEventHookThread(WinEventDelegate showCallback)
        {
            _showCallback = showCallback;
        }

        public void Start()
        {
            _thread = new Thread(Run)
            {
                IsBackground = true,
                Priority = ThreadPriority.Highest,
                Name = "WinTab Explorer WinEvent Hook"
            };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();

            if (!_started.Wait(2_000))
                DebugLog("WinEvent hook thread did not start within 2000ms");
        }

        private void Run()
        {
            try
            {
                _threadId = WinApi.GetCurrentThreadId();
                _ = WinApi.PeekMessage(out _, 0, 0, 0, WinApi.PM_NOREMOVE);

                const uint hookFlags = WinApi.WINEVENT_OUTOFCONTEXT | WinApi.WINEVENT_SKIPOWNPROCESS;
                _foregroundHookId = WinApi.SetWinEventHook(WinApi.EVENT_SYSTEM_FOREGROUND, WinApi.EVENT_SYSTEM_FOREGROUND, 0, _showCallback, 0, 0, hookFlags);
                _showHookId = WinApi.SetWinEventHook(WinApi.EVENT_OBJECT_CREATE, WinApi.EVENT_OBJECT_SHOW, 0, _showCallback, 0, 0, hookFlags);
                DebugLog($"WinEvent hook thread started id={_threadId} foregroundHook={_foregroundHookId} showHook={_showHookId}");
                _started.Set();

                while (WinApi.GetMessage(out var message, 0, 0, 0) > 0)
                {
                    WinApi.TranslateMessage(ref message);
                    WinApi.DispatchMessage(ref message);
                }
            }
            finally
            {
                if (_foregroundHookId != 0)
                    WinApi.UnhookWinEvent(_foregroundHookId);

                if (_showHookId != 0)
                    WinApi.UnhookWinEvent(_showHookId);

                DebugLog("WinEvent hook thread stopped");
                _stopped.Set();
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;

            if (_threadId != 0)
                WinApi.PostThreadMessage(_threadId, WinApi.WM_QUIT, 0, 0);

            if (!_stopped.Wait(2_000))
                DebugLog("WinEvent hook thread did not stop within 2000ms");

            _started.Dispose();
            _stopped.Dispose();
        }
    }
    private List<(InternetExplorer Window, WindowInfo WindowInfo)> AdoptNewShellWindows()
    {
        var result = new List<(InternetExplorer Window, WindowInfo WindowInfo)>();
        var singleTabTopLevelsInBatch = new HashSet<nint>();
        var count = _shellWindows.Count;

        for (var i = count - 1; i >= 0; i--)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window) continue;

            WindowInfo windowInfo;
            nint hWnd;
            bool wasTrackedTopLevel;
            string initialLocation = string.Empty;
            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.Keys.Contains(window)) continue;
                if (window.GetProperty("seenBefore") is not null) continue;

                hWnd = new IntPtr(window.HWND);
                wasTrackedTopLevel = HasTrackedTopLevelWindow(hWnd);
                var tabCount = Helper.GetAllExplorerTabs(hWnd).Take(2).Count();
                initialLocation = TryGetLocation(window);
                if (tabCount <= 1 &&
                    (wasTrackedTopLevel || !singleTabTopLevelsInBatch.Add(hWnd)))
                {
                    window.PutProperty("seenBefore", true);
                    continue;
                }

                window.PutProperty("seenBefore", true);
                if (!wasTrackedTopLevel &&
                    (_isForcingTabs || _reuseTabs) &&
                    !Helper.IsCtrlShiftDown() &&
                    _mainWindowHandle != hWnd)
                {
                    HideMergeSourceWindow(hWnd);
                }

                windowInfo = new WindowInfo();
                _windowEntryDict.Add(window, windowInfo);

                if (_windowEntryDict.Count == 1)
                {
                    _mainWindowHandle = hWnd;

                    if (SettingsManager.RestorePreviousWindows && _closedWindows.Any(w => w.Restore))
                        _ = RestorePreviousWindows();
                }
            }

            if (!wasTrackedTopLevel)
                TryHideRegisteredMergeSourceWindow(hWnd);
            result.Add((window, windowInfo));
        }

        return result;
    }
    private void OnShellWindowRegistered(int __)
    {
        // Keep the ShellWindows COM event callback non-blocking. Explorer can
        // raise this while it is committing shell-view edits such as renames,
        // and synchronous enumeration here can leave explorer.exe waiting on us.
        StartMergeSourceConcealPulse();
        ScheduleShellWindowRegistration(1);
    }
    private void ScheduleShellWindowRegistration(int delayMs = 25)
    {
        if (Interlocked.Exchange(ref _shellWindowRegistrationScheduled, 1) == 1)
            return;

        _ = Task.Run(async () =>
        {
            try
            {
                if (delayMs > 0)
                    await Task.Delay(delayMs);

                Interlocked.Exchange(ref _shellWindowRegistrationScheduled, 0);
                await ProcessRegisteredShellWindowsAsync();
            }
            catch
            {
                //
            }

            await Task.Delay(150);
            if (HasUntrackedShellWindows())
                ScheduleShellWindowRegistration(1);
        });
    }
    private bool HasUntrackedShellWindows()
    {
        if (_shellWindows == null)
            return false;

        try
        {
            var count = _shellWindows.Count;
            for (var i = count - 1; i >= 0; i--)
            {
                if (_shellWindows.Item(i) is not InternetExplorer window)
                    continue;

                lock (_windowEntryDictLock)
                {
                    if (!_windowEntryDict.Keys.Contains(window) &&
                        window.GetProperty("seenBefore") is null)
                    {
                        return true;
                    }
                }
            }
        }
        catch
        {
            //
        }

        return false;
    }
    private async Task ProcessRegisteredShellWindowsAsync()
    {
        if (_shellWindows == null)
            return;

        for (var i = 0; i < 8; i++)
        {
            await _shellWindowRegistrationLock.WaitAsync();
            List<(InternetExplorer Window, WindowInfo WindowInfo)> windows;
            try
            {
                windows = AdoptNewShellWindows();
            }
            finally
            {
                _shellWindowRegistrationLock.Release();
            }

            if (windows.Count == 0)
            {
                await Task.Delay(75);
                if (!HasUntrackedShellWindows())
                    return;

                continue;
            }

            await Task.WhenAll(windows.Select(item => ProcessRegisteredShellWindowAsync(item.Window, item.WindowInfo)));
        }
    }
    private async Task ProcessRegisteredShellWindowAsync(InternetExplorer window, WindowInfo windowInfo)
    {
        var showAgain = true;
        var removed = false;
        nint hWnd = 0;

        try
        {
            hWnd = new IntPtr(window.HWND);
            if (_processedHWnds.ContainsKey(hWnd))
                return;

            if (Helper.IsCtrlShiftDown())
            {
                DebugLog($"Registered release ctrl-shift hwnd={hWnd}");
                await RestoreMergeSourceWindowAsync(hWnd);
                RegisterIndependentWindow(window, windowInfo, hWnd);
                return;
            }

            if (HasOtherTrackedShellWindowForTopLevel(window, hWnd))
            {
                DebugLog($"Registered sibling hwnd={hWnd}");
                await PublishTabHandleForTrackedShellWindowAsync(window);
                RegisterIndependentWindow(window, windowInfo, hWnd);
                return;
            }

            var targetWindow = GetMainWindowHWnd(hWnd);
            if (!_isForcingTabs && !_reuseTabs)
            {
                DebugLog($"Registered release disabled hwnd={hWnd}");
                await RestoreMergeSourceWindowAsync(hWnd);
                RegisterIndependentWindow(window, windowInfo, hWnd);
                return;
            }

            if (_processedHWnds.ContainsKey(hWnd))
                return;

            HideMergeSourceWindow(hWnd);

            var location = await ResolveInitialLocation(window, hWnd);
            DebugLog($"Registered resolved hwnd={hWnd} target={targetWindow} location={location}");
            if (string.IsNullOrWhiteSpace(location) ||
                location.StartsWith("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", StringComparison.OrdinalIgnoreCase))
            {
                DebugLog($"Registered release unsupported-location hwnd={hWnd} location={location}");
                await RestoreMergeSourceWindowAsync(hWnd);
                RegisterIndependentWindow(window, windowInfo, hWnd);
                return;
            }

            var sourceAlive = Helper.IsFileExplorerWindow(hWnd);
            if (sourceAlive && !IsStartupExplorerLocation(location))
            {
                _ = await GetTabHandle(window);
                var tabCount = await WaitForExplorerTabCount(hWnd);
                if (tabCount != 1)
                {
                    DebugLog($"Registered release tab-count hwnd={hWnd} count={tabCount}");
                    await RestoreMergeSourceWindowAsync(hWnd);
                    RegisterIndependentWindow(window, windowInfo, hWnd);
                    return;
                }
            }

            targetWindow = GetMainWindowHWnd(hWnd, location);
            if (targetWindow == 0 || hWnd == targetWindow)
            {
                DebugLog($"Registered release no-target hwnd={hWnd} location={location}");
                await RestoreMergeSourceWindowAsync(hWnd);
                RegisterIndependentWindow(window, windowInfo, hWnd);
                return;
            }

            if (sourceAlive && TryGetRecentlyClosedWindow(location, out var closedWindow))
            {
                DebugLog($"Registered release recently-closed hwnd={hWnd} location={location}");
                SelectItems(window, closedWindow!.SelectedItems);
                await RestoreMergeSourceWindowAsync(hWnd);
                RegisterIndependentWindow(window, windowInfo, hWnd);
                return;
            }

            if (sourceAlive && !IsStartupExplorerLocation(location))
                HideMergeSourceWindow(hWnd);

            var record = new WindowRecord(location, hWnd, TryGetSelectedItems(window));
            if (!await OpenTabNavigateWithSelection(record, targetWindow))
            {
                DebugLog($"Registered merge-failed hwnd={hWnd} location={location}");
                if (Helper.IsFileExplorerWindow(hWnd))
                {
                    await RestoreMergeSourceWindowAsync(hWnd);
                    RegisterIndependentWindow(window, windowInfo, hWnd);
                }
                else
                {
                    DebugLog($"Registered merge-failed source already closed; not reopening intermediate hwnd={hWnd} location={location}");
                    RemoveMergeSourceTracking(hWnd);
                    RemoveWindowAndUnhookEvents(window, windowInfo, restoreHiddenWindow: false);
                    removed = true;
                }
                return;
            }

            showAgain = false;
            DebugLog($"Registered merge-succeeded hwnd={hWnd} location={location}");
            UnhookWindowEvents(window, windowInfo);
            if (await CloseMergedSourceWindowAsync(window, hWnd))
            {
                RemoveMergeSourceTracking(hWnd);
                RemoveWindowAndUnhookEvents(window, windowInfo, restoreHiddenWindow: false);
                removed = true;
            }
            else
            {
                RegisterIndependentWindow(window, windowInfo, hWnd);
            }
        }
        catch (Exception ex)
        {
            DebugLog($"Registered error hwnd={hWnd} error={ex.GetType().Name}:{ex.Message}");
        }
        finally
        {
            if (!removed && showAgain && hWnd != 0)
            {
                await RestoreMergeSourceWindowAsync(hWnd);
                RegisterIndependentWindow(window, windowInfo, hWnd);
            }
        }
    }
    private void RegisterIndependentWindow(InternetExplorer window, WindowInfo windowInfo, nint hWnd)
    {
        if (hWnd != 0)
            PreventWindowHiding(hWnd);

        HookWindowEvents(window, windowInfo);
    }
    private async Task PublishTabHandleForTrackedShellWindowAsync(InternetExplorer window)
    {
        _ = await GetTabHandle(window);
    }
    private async Task<bool> CloseMergedSourceWindowAsync(InternetExplorer window, nint hWnd)
    {
        RequestCloseMergedSourceWindow(window, hWnd);
        if (hWnd == 0)
            return true;

        var closed = await Helper.DoUntilConditionAsync(
            () => !Helper.IsFileExplorerWindow(hWnd),
            isClosed => isClosed,
            1_500,
            50);

        if (closed)
            return true;

        RequestCloseMergedSourceWindow(window, hWnd);
        closed = await Helper.DoUntilConditionAsync(
            () => !Helper.IsFileExplorerWindow(hWnd),
            isClosed => isClosed,
            900,
            50);

        if (!closed)
            await RestoreMergeSourceWindowAsync(hWnd);

        return closed;
    }
    private static void RequestCloseMergedSourceWindow(InternetExplorer window, nint hWnd)
    {
        if (hWnd != 0)
            WinApi.PostMessage(hWnd, WinApi.WM_CLOSE, 0, 0);

        _ = Task.Run(() =>
        {
            try
            {
                window.Quit();
            }
            catch
            {
                //
            }
        });
    }
    private Task<string> ResolveInitialLocation(InternetExplorer window, nint hWnd = 0)
    {
        return _locationResolver.ResolveAsync(
            () => TryGetLocation(window),
            IsStartupExplorerLocation,
            isBusy: () => IsShellWindowBusy(window));
    }
    private static bool IsShellWindowBusy(InternetExplorer window)
    {
        try
        {
            return window.Busy;
        }
        catch
        {
            return false;
        }
    }
    private string TryGetLocation(InternetExplorer window)
    {
        try
        {
            return GetLocation(window);
        }
        catch
        {
            return string.Empty;
        }
    }
    private static nint SafeGetWindowHandle(InternetExplorer window)
    {
        try
        {
            return new IntPtr(window.HWND);
        }
        catch
        {
            return 0;
        }
    }
    private bool IsStartupExplorerLocation(string location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return true;

        location = Helper.NormalizeLocation(location);
        if (StringComparer.OrdinalIgnoreCase.Equals(location, _defaultLocation))
            return true;

        if (_startupLocationCache.TryGetValue(location, out var cached))
            return cached;

        bool isStartup;
        try
        {
            isStartup = _shellPathComparer.IsEquivalent(location, _defaultLocation);
        }
        catch
        {
            isStartup = false;
        }

        if (_startupLocationCache.Count >= StartupLocationCacheLimit)
            _startupLocationCache.Clear();

        _startupLocationCache[location] = isStartup;
        return isStartup;
    }
    private static Task<int> WaitForExplorerTabCount(nint hWnd)
    {
        return Helper.DoUntilConditionAsync(
            () => Helper.GetAllExplorerTabs(hWnd).Take(2).Count(),
            count => count > 0,
            1_500,
            25);
    }
    private static async Task RestoreHiddenExplorerWindowAsync(nint hWnd)
    {
        await Helper.DoUntilConditionAsync(
            () => Helper.RestoreHiddenExplorerWindow(hWnd, removeCache: false, removeLayeredStyle: !SettingsManager.HaveThemeIssue),
            restored => restored || !Helper.HiddenWindows.ContainsKey(hWnd),
            1_500,
            50);

        // OnWindowShown can arrive after a release decision. Keep the cache briefly,
        // then allow later show events to be evaluated normally.
        _ = Task.Delay(3_000).ContinueWith(t => Helper.HiddenWindows.TryRemove(hWnd, out _), TaskScheduler.Default);
    }
    private void HookWindowEvents(InternetExplorer window, WindowInfo windowInfo)
    {
        if (windowInfo.EventsHooked)
            return;

        var hookedTopLevelHWnd = SafeGetWindowHandle(window);
        if (hookedTopLevelHWnd != 0)
        {
            windowInfo.HookedTopLevelHWnd = hookedTopLevelHWnd;
            _hookedTopLevelUseCounts.AddOrUpdate(hookedTopLevelHWnd, 1, (_, count) => count + 1);
        }

        // Create strongly-typed handlers so we can remove them later
        windowInfo.OnQuitHandler = () =>
        {
            var location = windowInfo.Location ?? GetLocation(window);
            var locationName = windowInfo.Name ?? window.LocationName;
            var windowRecord = new WindowRecord(location, new IntPtr(window.HWND), name: locationName);
            lock (_closedWindowsLock) _closedWindows.Add(windowRecord);

            // Home, This PC, etc
            if (location == _defaultLocation)
            {
                RemoveWindowAndUnhookEvents(window, windowInfo);
                return;
            }

            windowRecord.SelectedItems = GetSelectedItems(window);
            RemoveWindowAndUnhookEvents(window, windowInfo);
        };

        if (SettingsManager.RestorePreviousWindows)
            windowInfo.OnNavigateHandler = (object _, ref object _) =>
            {
                windowInfo.Location = GetLocation(window);
                windowInfo.Name = window.LocationName;
            };

        try
        {
            // Subscribe
            window.OnQuit += windowInfo.OnQuitHandler;
            if (SettingsManager.RestorePreviousWindows)
            {
                windowInfo.Location = GetLocation(window);
                windowInfo.Name = window.LocationName;
                window.NavigateComplete2 += windowInfo.OnNavigateHandler;
            }

            windowInfo.EventsHooked = true;

            // Make sure the window is still alive (User might have closed it immediately after opening it)
            var hWnd = new IntPtr(window.HWND);
            if (Helper.IsFileExplorerWindow(hWnd))
                ScheduleStripBoundsRefresh(hWnd);
        }
        catch
        {
            ReleaseHookedTopLevel(windowInfo);
            lock (_windowEntryDictLock)
                _windowEntryDict.Remove(window);
        }
    }
    private void UnhookWindowEvents(InternetExplorer window, WindowInfo windowInfo)
    {
        if (!windowInfo.EventsHooked)
            return;

        if (windowInfo.OnQuitHandler != null) window.OnQuit -= windowInfo.OnQuitHandler;
        if (windowInfo.OnNavigateHandler != null) window.NavigateComplete2 -= windowInfo.OnNavigateHandler;
        windowInfo.EventsHooked = false;
        ReleaseHookedTopLevel(windowInfo);
    }
    private void ReleaseHookedTopLevel(WindowInfo windowInfo)
    {
        var hWnd = windowInfo.HookedTopLevelHWnd;
        if (hWnd == 0)
            return;

        _hookedTopLevelUseCounts.AddOrUpdate(hWnd, 0, (_, count) => Math.Max(0, count - 1));
        if (_hookedTopLevelUseCounts.TryGetValue(hWnd, out var remaining) && remaining <= 0)
            _hookedTopLevelUseCounts.TryRemove(hWnd, out _);

        windowInfo.HookedTopLevelHWnd = 0;
    }
    private void RemoveWindowAndUnhookEvents(InternetExplorer window, WindowInfo windowInfo, bool useLock = true, bool restoreHiddenWindow = true)
    {
        UnhookWindowEvents(window, windowInfo);

        // Remove from dictionary
        if (useLock)
        {
            lock (_windowEntryDictLock)
                _windowEntryDict.Remove(window);
        }
        else
            _windowEntryDict.Remove(window);

        try
        {
            var hWnd = new IntPtr(window.HWND);
            if (restoreHiddenWindow)
                Helper.RestoreHiddenExplorerWindow(hWnd, removeCache: true, removeLayeredStyle: !SettingsManager.HaveThemeIssue);

            _processedHWnds.TryRemove(hWnd, out _);
            ReleaseTopLevelTrackingIfUnused(hWnd);
            InvalidateAutomationRoot(hWnd);
            InvalidateStripBounds(hWnd);
            if (_mainWindowHandle == hWnd && !Helper.IsFileExplorerWindow(hWnd))
                _mainWindowHandle = 0;
        }
        catch
        {
            //
        }

        // Finally, release the COM reference for this InternetExplorer instance
        Marshal.ReleaseComObject(window);
    }

    private async Task RestorePreviousWindows()
    {
        var result = await RunInStaThread(() => CustomMessageBox.Show(
            "Do you want to restore previously opened windows?",
            Constants.AppName,
            MessageBoxButton.YesNo,
            MessageBoxImage.Question));

        foreach (var record in _closedWindows.Where(record => record.Restore))
        {
            record.Restore = false;
            
            if (result != MessageBoxResult.Yes) continue;
            
            _ = OpenTabNavigateWithSelection(record);
        }
    }
    private async Task OpenNewWindowWithSelection(WindowRecord windowToOpen, bool duplicate = true, bool lockToOpenWindows = true)
    {
        if (lockToOpenWindows)
            await _toOpenWindowsLock.WaitAsync();

        try
        {
            lock (_closedWindowsLock)
                _closedWindows.Add(windowToOpen);

            var hasSelection = windowToOpen.SelectedItems?.Length > 0;

            nint[]? currentWindows = null;
            if (hasSelection)
                currentWindows = Helper.GetAllExplorerWindows().ToArray();

            Helper.BypassWinForegroundRestrictions();

            var location = string.IsNullOrWhiteSpace(windowToOpen.Location) ? _defaultLocation : windowToOpen.Location;
            await RunInStaThread(() =>
            {
                Shell? shell = null;
                try
                {
                    shell = new Shell();
                    shell.ShellExecute(location, "", "", duplicate ? "opennewwindow" : "open");
                }
                finally
                {
                    if (shell != null)
                        Marshal.ReleaseComObject(shell);
                }
            });

            if (!hasSelection) return;

            var newWindowHandle = await Helper.ListenForNewExplorerWindowAsync(currentWindows ?? []);
            if (newWindowHandle == 0) return;

            var window = _windowEntryDict.Keys.FirstOrDefault(w => w.HWND == newWindowHandle);
            if (window == null) return;

            SelectItems(window, windowToOpen.SelectedItems);
        }
        finally
        {
            if (lockToOpenWindows)
                _toOpenWindowsLock.Release();
        }
    }
    private async Task<bool> OpenTabNavigateWithSelection(
        WindowRecord windowToOpen,
        nint windowHandle = 0,
        bool isDuplicate = false,
        bool forceTabReuse = false)
    {
        DebugLog($"OpenTab begin target={windowToOpen.Location}");
        nint mainWindowHWnd = 0;
        nint newTabHandle = 0;
        InternetExplorer? window = null;

        await _toOpenWindowsLock.WaitAsync();
        try
        {
            DebugLog($"OpenTab lock target={windowToOpen.Location}");
            if ((_reuseTabs || forceTabReuse) && !isDuplicate && _windowEntryDict.Count > 0 && !string.IsNullOrWhiteSpace(windowToOpen.Location))
            {
                if (TrySearchForTab(windowToOpen.Location, windowToOpen.Handle, out var existingTab, out var existingWindow))
                {
                    windowHandle = WinApi.GetParent(existingTab);
                    WinApi.RestoreWindowToForeground(windowHandle);
                    if (await SelectTabByUniqueNameVerified(windowHandle, existingTab, 700, existingWindow))
                    {
                        DebugLog($"OpenTab reused target={windowToOpen.Location}");
                        return true;
                    }

                    DebugLog($"OpenTab reuse-select-failed target={windowToOpen.Location}");
                    WinApi.RestoreWindowToForeground(windowHandle);
                    return true;
                }
            }

            // Get the main window
            mainWindowHWnd = Helper.IsFileExplorerWindow(windowHandle)
                ? windowHandle
                : GetMainWindowHWnd(windowToOpen.Handle);

            if (mainWindowHWnd == 0)
            {
                await OpenNewWindowWithSelection(windowToOpen, lockToOpenWindows: false);
                DebugLog($"OpenTab opened-window target={windowToOpen.Location}");
                return true;
            }

            try
            {
                var currentTabs = Helper.GetAllExplorerTabs(mainWindowHWnd).ToArray();
                DebugLog($"OpenTab main={mainWindowHWnd} tabs={currentTabs.Length} target={windowToOpen.Location}");

                await RequestToOpenNewTab(mainWindowHWnd, lockToOpenWindows: false);
                DebugLog($"OpenTab requested main={mainWindowHWnd} target={windowToOpen.Location}");

                newTabHandle = await Helper.ListenForNewExplorerTabAsync(mainWindowHWnd, currentTabs, 2_000);
                if (newTabHandle == 0)
                {
                    DebugLog($"OpenTab no-new-tab target={windowToOpen.Location}");
                    return false;
                }
                DebugLog($"OpenTab new-tab={newTabHandle} target={windowToOpen.Location}");

                window = await Helper.DoUntilNotDefaultAsync(
                    () => FindShellWindowByTabHandle(newTabHandle, mainWindowHWnd),
                    2_000,
                    50);

                if (window == null)
                {
                    await CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle);
                    DebugLog($"OpenTab missing-shell-window tab={newTabHandle} target={windowToOpen.Location}");
                    return false;
                }
                DebugLog($"OpenTab found-shell-window tab={newTabHandle} target={windowToOpen.Location}");
            }
            catch (Exception ex)
            {
                await CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle);
                DebugLog($"OpenTab error tab={newTabHandle} target={windowToOpen.Location} error={ex.GetType().Name}:{ex.Message}");
                return false;
            }
        }
        finally
        {
            _toOpenWindowsLock.Release();
        }

        try
        {
            if (window == null)
                return false;

            if (!await NavigateNewTabToTargetAsync(window, windowToOpen.Location))
            {
                await CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle);
                return false;
            }

            WinApi.RestoreWindowToForeground(mainWindowHWnd);
            SelectItems(window, windowToOpen.SelectedItems);
            return true;
        }
        catch (Exception ex)
        {
            await CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle);
            DebugLog($"OpenTab post-create error tab={newTabHandle} target={windowToOpen.Location} error={ex.GetType().Name}:{ex.Message}");
            return false;
        }
    }
    private static async Task CloseFailedNewTabAsync(nint parentWindowHandle, nint tabHandle)
    {
        TryCloseFailedNewTab(tabHandle);

        var stillExists = false;
        if (parentWindowHandle != 0 && tabHandle != 0)
        {
            stillExists = await Helper.DoUntilConditionAsync(
                () => Helper.GetAllExplorerTabs(parentWindowHandle).Contains(tabHandle),
                exists => !exists,
                1_200,
                50);
        }

        if (stillExists)
        {
            TryCloseFailedNewTab(tabHandle);
            if (parentWindowHandle != 0 && tabHandle != 0)
            {
                await Helper.DoUntilConditionAsync(
                    () => Helper.GetAllExplorerTabs(parentWindowHandle).Contains(tabHandle),
                    exists => !exists,
                    800,
                    50);
            }
        }

        await Task.Delay(100);
    }
    private static void TryCloseFailedNewTab(nint tabHandle)
    {
        if (tabHandle == 0)
            return;

        try
        {
            // Send 0xA021 magic command (CTRL + W)
            WinApi.SendMessage(tabHandle, WinApi.WM_COMMAND, 0xA021, 1);
        }
        catch
        {
            //
        }
    }
    private async Task<bool> WaitForNavigation(InternetExplorer window, string targetLocation, int timeoutMs = 5_000)
    {
        if (string.IsNullOrWhiteSpace(targetLocation))
            return true;

        var resolvedLocation = await Helper.DoUntilConditionAsync(
            () => TryGetLocation(window),
            location => AreLocationsEquivalent(location, targetLocation),
            timeoutMs,
            50);

        return AreLocationsEquivalent(resolvedLocation, targetLocation);
    }
    private async Task<bool> NavigateNewTabToTargetAsync(InternetExplorer window, string targetLocation)
    {
        if (string.IsNullOrWhiteSpace(targetLocation))
            return true;

        if (AreLocationsEquivalent(TryGetLocation(window), targetLocation))
            return true;

        var navigationCompleted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        DWebBrowserEvents2_NavigateComplete2EventHandler? navigateHandler = null;
        navigateHandler = (object _, ref object _) =>
        {
            if (AreLocationsEquivalent(TryGetLocation(window), targetLocation))
                navigationCompleted.TrySetResult(true);
        };

        try
        {
            window.NavigateComplete2 += navigateHandler;
            if (!await NavigateToTargetIfNeeded(window, targetLocation))
            {
                DebugLog($"OpenTab navigate-failed target={targetLocation}");
                return false;
            }

            DebugLog($"OpenTab navigated target={targetLocation}");
            var completed = await Task.WhenAny(navigationCompleted.Task, Task.Delay(NavigationCompleteWaitMs));
            if (completed == navigationCompleted.Task && await navigationCompleted.Task)
                return true;

            if (await WaitForNavigation(window, targetLocation, NavigationVerificationWaitMs))
                return true;

            if (AreLocationsEquivalent(TryGetLocation(window), targetLocation))
                return true;

            DebugLog($"OpenTab target-check-failed target={targetLocation} current={TryGetLocation(window)}");
            return false;
        }
        finally
        {
            if (navigateHandler != null)
                window.NavigateComplete2 -= navigateHandler;
        }
    }
    private async Task<bool> NavigateToTargetIfNeeded(InternetExplorer window, string targetLocation)
    {
        if (string.IsNullOrWhiteSpace(targetLocation))
            return true;

        try
        {
            if (AreLocationsEquivalent(TryGetLocation(window), targetLocation))
                return true;

            await Navigate(window, targetLocation);
            return true;
        }
        catch
        {
            return false;
        }
    }
    private bool AreLocationsEquivalent(string location, string targetLocation)
    {
        if (string.IsNullOrWhiteSpace(location) || string.IsNullOrWhiteSpace(targetLocation))
            return false;

        if (StringComparer.OrdinalIgnoreCase.Equals(location, targetLocation))
            return true;

        var normalizedLocation = Helper.NormalizeLocation(location);
        var normalizedTargetLocation = Helper.NormalizeLocation(targetLocation);
        if (StringComparer.OrdinalIgnoreCase.Equals(normalizedLocation, normalizedTargetLocation))
            return true;

        try
        {
            return _shellPathComparer.IsEquivalent(location, targetLocation);
        }
        catch
        {
            return false;
        }
    }
    private bool TryGetRecentlyClosedWindow(string location, out WindowRecord? closedWindow, int maxAge = 2_000)
    {
        if (string.IsNullOrWhiteSpace(location))
        {
            closedWindow = null;
            return false;
        }

        nint targetPidl = 0;
        try
        {
            targetPidl = _shellPathComparer.GetPidlFromPath(location);
            lock (_closedWindowsLock)
            {
                for (var i = _closedWindows.Count - 1; i >= 0; i--)
                {
                    var record = _closedWindows[i];
                    if (Environment.TickCount - record.CreatedAt > maxAge) break;
                    if (!_shellPathComparer.IsEquivalent(location, record.Location, targetPidl)) continue;
                    _closedWindows.RemoveAt(i);
                    closedWindow = record;
                    return true;
                }
            }
            closedWindow = null;
            return false;
        }
        finally
        {
            if (targetPidl != 0)
                Marshal.FreeCoTaskMem(targetPidl);
        }
    }
    private nint GetMainWindowHWnd(nint otherThan, string? targetLocation = null)
    {
        var preferNonStartupTarget = ShouldPreferNonStartupTarget(targetLocation);

        if (Helper.IsFileExplorerForeground(out var foregroundWindow) &&
            IsPreferredMergeTargetWindow(foregroundWindow, otherThan, targetLocation))
        {
            _mainWindowHandle = foregroundWindow;
            return _mainWindowHandle;
        }

        if (IsPreferredMergeTargetWindow(_mainWindowHandle, otherThan, targetLocation))
            return _mainWindowHandle;

        var allWindows = WinApi.FindAllWindowsEx("CabinetWClass").ToArray();
        var tabCounts = new Dictionary<nint, int>();

        int GetCachedTabCount(nint hWnd)
        {
            if (tabCounts.TryGetValue(hWnd, out var count))
                return count;

            count = WinApi.FindAllWindowsEx("ShellTabWindowClass", hWnd).Count();
            tabCounts[hWnd] = count;
            return count;
        }

        nint SelectBestMergeTarget(Func<nint, bool> predicate)
        {
            nint bestWindow = 0;
            var bestTabCount = -1;

            // Iterate from the end to keep the previous "oldest window wins ties" behavior.
            for (var i = allWindows.Length - 1; i >= 0; i--)
            {
                var hWnd = allWindows[i];
                if (!predicate(hWnd))
                    continue;

                var tabCount = GetCachedTabCount(hWnd);
                if (tabCount <= bestTabCount)
                    continue;

                bestTabCount = tabCount;
                bestWindow = hWnd;
            }

            return bestWindow;
        }

        // Get another handle other than the newly created one. (In case if it is still alive.)
        _mainWindowHandle = SelectBestMergeTarget(h => IsPreferredMergeTargetWindow(h, otherThan, targetLocation));

        if (_mainWindowHandle != 0) return _mainWindowHandle;

        if (preferNonStartupTarget)
        {
            _mainWindowHandle = SelectBestMergeTarget(h => IsStableMergeTargetWindow(h, otherThan));

            if (_mainWindowHandle != 0) return _mainWindowHandle;
        }

        _mainWindowHandle = SelectBestMergeTarget(h => IsFallbackMergeTargetWindow(h, otherThan));

        return _mainWindowHandle;
    }
    private bool IsPreferredMergeTargetWindow(nint hWnd, nint otherThan, string? targetLocation)
    {
        if (!IsStableMergeTargetWindow(hWnd, otherThan))
            return false;

        return !ShouldPreferNonStartupTarget(targetLocation) ||
               HasNonStartupShellWindowForTopLevel(hWnd);
    }
    private bool ShouldPreferNonStartupTarget(string? targetLocation)
    {
        return !string.IsNullOrWhiteSpace(targetLocation) &&
               !IsStartupExplorerLocation(targetLocation);
    }
    private bool HasNonStartupShellWindowForTopLevel(nint hWnd)
    {
        lock (_windowEntryDictLock)
        {
            foreach (var (window, info, _) in _windowEntryDict)
            {
                try
                {
                    if (new IntPtr(window.HWND) != hWnd)
                        continue;

                    var location = info.Location ?? TryGetLocation(window);
                    if (!IsStartupExplorerLocation(location))
                        return true;
                }
                catch
                {
                    //
                }
            }
        }

        return false;
    }
    private bool IsStableMergeTargetWindow(nint hWnd, nint otherThan)
    {
        if (hWnd == 0 || hWnd == otherThan)
            return false;
        if (!Helper.IsFileExplorerWindow(hWnd))
            return false;
        if (!HasHookedShellWindowForTopLevel(hWnd))
            return false;

        return GetActiveTabHandle(hWnd) != 0;
    }
    private static bool IsFallbackMergeTargetWindow(nint hWnd, nint otherThan)
    {
        if (hWnd == 0 || hWnd == otherThan)
            return false;
        if (!Helper.IsFileExplorerWindow(hWnd))
            return false;

        return GetActiveTabHandle(hWnd) != 0;
    }
    private Task<nint> GetTabHandle(InternetExplorer window)
    {
        if (_windowEntryDict.TryGetValue(window, out WindowEntry entry) && entry.OptionalKey is { } handle and > 0)
            return Task.FromResult(handle);

        return QueryTabHandle(window, updateDictionary: true);
    }
    private Task<nint> QueryTabHandle(InternetExplorer window, bool updateDictionary)
    {
        return RunInStaThread(() =>
        {
            // ReSharper disable once SuspiciousTypeConversion.Global
            if (window is not Interop.IServiceProvider sp) return 0;

            sp.QueryService(ref _shellBrowserGuid, ref _shellBrowserGuid, out var shellBrowser);
            if (shellBrowser == null) return 0;

            try
            {
                shellBrowser.GetWindow(out var hWnd);

                if (updateDictionary && hWnd != 0)
                {
                    lock (_windowEntryDictLock)
                    {
                        if (_windowEntryDict.ContainsKey(window))
                            _windowEntryDict.UpdateOptionalKey(window, hWnd);
                    }

                }

                return hWnd;
            }
            finally
            {
                Marshal.ReleaseComObject(shellBrowser);
            }
        });
    }
    private async Task<InternetExplorer?> FindShellWindowByTabHandle(nint tabHandle, nint parentWindowHandle = 0)
    {
        var cachedWindow = GetWindowByTabHandle(tabHandle, parentWindowHandle);
        if (cachedWindow != null)
            return cachedWindow;

        var count = _shellWindows.Count;
        for (var i = count - 1; i >= 0; i--)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;

            if (parentWindowHandle != 0)
            {
                try
                {
                    if (new IntPtr(window.HWND) != parentWindowHandle)
                        continue;
                }
                catch
                {
                    continue;
                }
            }

            var currentTabHandle = await QueryTabHandle(window, updateDictionary: false);
            if (currentTabHandle != tabHandle)
                continue;

            WindowInfo windowInfo;
            InternetExplorer windowToReturn = window;
            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.TryGetValue(tabHandle, out InternetExplorer? existingWindow) && existingWindow != null)
                {
                    windowToReturn = existingWindow;
                    if (!_windowEntryDict.TryGetValue(windowToReturn, out windowInfo!))
                        continue;
                }
                else if (_windowEntryDict.TryGetValue(window, out windowInfo!))
                {
                    _windowEntryDict.UpdateOptionalKey(window, tabHandle);
                }
                else
                {
                    window.PutProperty("seenBefore", true);
                    windowInfo = new WindowInfo();
                    _windowEntryDict.Add(window, windowInfo, tabHandle);
                }
            }

            HookWindowEvents(windowToReturn, windowInfo);
            return windowToReturn;
        }

        return null;
    }
    private static nint GetActiveTabHandle(nint windowHandle)
    {
        // Active tab always at the top of the z-index
        return WinApi.FindWindowEx(windowHandle, 0, "ShellTabWindowClass", null);
    }
    private InternetExplorer? GetWindowByTabHandle(nint tabHandle)
    {
        if (tabHandle == 0) return null;
        return _windowEntryDict.TryGetValue(tabHandle, out InternetExplorer? foundWindow) ? foundWindow : null;
    }
    private InternetExplorer? GetWindowByTabHandle(nint tabHandle, nint parentWindowHandle)
    {
        var window = GetWindowByTabHandle(tabHandle);
        if (window == null)
            return null;

        if (parentWindowHandle == 0)
            return window;

        try
        {
            return new IntPtr(window.HWND) == parentWindowHandle ? window : null;
        }
        catch
        {
            return null;
        }
    }
    private static string[]? GetSelectedItems(InternetExplorer window)
    {
        var selectedItems = (window.Document as ShellFolderView)!.SelectedItems();
        var count = selectedItems.Count;
        if (count == 0) return null;

        var result = new string[count];
        for (var i = 0; i < count; i++)
        {
            result[i] = selectedItems.Item(i).Name;
        }

        return result;
    }
    private static string[]? TryGetSelectedItems(InternetExplorer window)
    {
        try
        {
            return GetSelectedItems(window);
        }
        catch
        {
            return null;
        }
    }
    private static void SelectItems(InternetExplorer window, string[]? names)
    {
        if (names == null || names.Length == 0) return;

        if (window.Document is not ShellFolderView document) return;

        for (var i = 0; i < names.Length; i++)
        {
            var name = names[i];
            object item = document.Folder.ParseName(name);
            if (item == null) continue;
            document.SelectItem(ref item, 1);
        }
    }
    private static string GetLocation(InternetExplorer window)
    {
        var path = window.LocationURL;
        if (!string.IsNullOrWhiteSpace(path)) return Helper.NormalizeLocation(path);

        // Recycle Bin, This PC, etc
        path = ((window.Document as ShellFolderView)!.Folder as Folder2)!.Self.Path;
        return Helper.NormalizeLocation(path);
    }
    private async Task Navigate(InternetExplorer window, string path)
    {
        if (!path.Contains("#") && !path.Contains("%23"))
        {
            window.Navigate2(path);
            return;
        }

        var folder = await RunInStaThread(() =>
        {
            Shell? shell = null;
            Folder? folder;
            try
            {
                shell = new Shell();
                folder = shell.NameSpace(path);
            }
            finally
            {
                if (shell != null)
                    Marshal.ReleaseComObject(shell);
            }
            return folder;
        });

        try
        {
            window.Navigate2(folder);
        }
        finally
        {
            if (folder != null)
                Marshal.ReleaseComObject(folder);
        }
    }
    private Task RunInStaThread(Action action, TaskCreationOptions tco = default, CancellationToken ct = default)
    {
        return Task.Factory.StartNew(action, ct, tco, _staTaskScheduler);
    }
    private Task<T?> RunInStaThread<T>(Func<T?> action, TaskCreationOptions tco = default, CancellationToken ct = default)
    {
        return Task.Factory.StartNew(action, ct, tco, _staTaskScheduler);
    }
    
    private void StartExplorerProcessCheck() => _explorerCheckTimer = new Timer(CheckForMainExplorer, null, 0, 1000);
    private void CheckForMainExplorer(object? state)
    {
        var process = Helper.GetMainExplorerProcess();
        if (process == null) return;
        
        _explorerCheckTimer?.Dispose();
        _explorerCheckTimer = null;
        
        lock (_processLock)
        {
            if (_mainExplorerProcessId != 0) return;
            
            _mainExplorerProcessId = process.Id;
            InitializeShellObjects();
            OnShellInitialized?.Invoke();
        }
    }
    private void OnExplorerProcessTerminated(object? s, ProcessEventArgs e)
    {
        // Main explorer.exe process (_shellWindows must be restarted)
        lock (_processLock)
        {
            if (e.ProcessId == _mainExplorerProcessId)
            {
                _mainExplorerProcessId = 0;
                DisposeShellObjects();
                StartExplorerProcessCheck();
                return;
            }
        }
        
        // Other explorer.exe processes
        lock (_windowEntryDictLock)
        {
            if (_windowEntryDict.Count == 0) return;
            var crashCount = 0;
            for (var i = _windowEntryDict.Count - 1; i >= 0; i--)
            {
                var (window, info) = _windowEntryDict.ElementAt<WindowEntry>(i);
                try
                {
                    _ = window.HWND;
                }
                catch
                {
                    if (info.OnNavigateHandler != null)
                    {
                        crashCount++;
                        lock (_closedWindowsLock)
                            _closedWindows.Add(new WindowRecord(info.Location!, name: info.Name!));
                    }
                    
                    RemoveWindowAndUnhookEvents(window, info, useLock: false);
                }
            }
            if (!SettingsManager.RestorePreviousWindows || _windowEntryDict.Count > 0) return;
            lock (_closedWindowsLock)
            {
                for (var i = 1; i <= crashCount; i++)
                    _closedWindows[_closedWindows.Count - i].Restore = true;
            }
        }
    }

    private void InitializeShellObjects()
    {
        _shellPathComparer = new ShellPathComparer();
        _staTaskScheduler = new StaTaskScheduler();
        _shellWindows = new ShellWindows();

        _defaultLocation = Helper.GetDefaultExplorerLocation(_shellPathComparer);
        ClearShellCaches();
        RecoverHiddenExplorerWindows("initialize-shell");

        if (Helper.IsFileExplorerForeground(out var foregroundWindow))
            _mainWindowHandle = foregroundWindow;
        
        if (SettingsManager.ClosedWindows != null)
            lock (_closedWindowsLock) _closedWindows.AddRange(SettingsManager.ClosedWindows);

        // Hook the global "WindowRegistered" event
        _windowRegisteredHandler = OnShellWindowRegistered;
        _shellWindows.WindowRegistered += _windowRegisteredHandler;

        // WinEvent only wakes ShellWindows processing; WindowRegistered owns merge and release.
        _eventObjectShowHookCallback = OnWindowShown;
        _winEventHookThread = new WinEventHookThread(_eventObjectShowHookCallback);
        _winEventHookThread.Start();

        // Hook the event handlers for already-open windows
        var hasOpen = false;
        var count = _shellWindows.Count;
        for (var i = 0; i < count; i++)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;
            if (_windowEntryDict.Keys.Contains(window))
                continue;
            if (window.GetProperty("seenBefore") is not null)
                continue;

            hasOpen = true;

            var windowInfo = new WindowInfo();
            _windowEntryDict.Add(window, windowInfo);
            window.PutProperty("seenBefore", true);
            PreventWindowHiding(new IntPtr(window.HWND));

            if (_mainWindowHandle == 0)
                _mainWindowHandle = new IntPtr(window.HWND);

            _ = GetTabHandle(window);
            HookWindowEvents(window, windowInfo);
        }

        if (!hasOpen) return;
        lock (_closedWindowsLock)
            foreach (var window in _closedWindows) window.Restore = false;
    }
    private void DisposeShellObjects()
    {
        RecoverHiddenExplorerWindows("dispose-shell");

        if (_shellWindows == null)
            return;

        PersistWindows();

        // Unhook global event
        if (_windowRegisteredHandler != null)
        {
            _shellWindows.WindowRegistered -= _windowRegisteredHandler;
            _windowRegisteredHandler = null;
        }
        _winEventHookThread?.Dispose();
        _winEventHookThread = null;
        _eventObjectShowHookCallback = null;

        List<WindowEntry> windowEntries;
        lock (_windowEntryDictLock)
            windowEntries = ((IEnumerable<WindowEntry>)_windowEntryDict).ToList();

        // Unsubscribe from each InternetExplorer instance's events
        foreach (var (window, windowInfo) in windowEntries)
        {
            // Unsubscribe
            if (windowInfo.OnQuitHandler != null) window.OnQuit -= windowInfo.OnQuitHandler;
            if (windowInfo.OnNavigateHandler != null) window.NavigateComplete2 -= windowInfo.OnNavigateHandler;

            // Release the COM object
            Marshal.ReleaseComObject(window);
        }
        lock (_windowEntryDictLock)
            _windowEntryDict.Clear();

        // Release the ShellWindows COM object
        Marshal.ReleaseComObject(_shellWindows);

        _shellPathComparer.Dispose();
        _staTaskScheduler.Dispose();
        ClearShellCaches();
        _shellWindows = null!;
        _shellPathComparer = null!;
        _staTaskScheduler = null!;
        _mainWindowHandle = 0;
    }

    private void ClearShellCaches()
    {
        _startupLocationCache.Clear();
        _automationRootCache.Clear();
        _stripBoundsCache.Clear();
        _stripBoundsRefreshInflight.Clear();
        _hookedTopLevelUseCounts.Clear();
    }

    private void RecoverHiddenExplorerWindows(string reason)
    {
        var candidates = Helper.GetAllExplorerWindows()
            .Concat(Helper.HiddenWindows.Keys)
            .Distinct()
            .ToArray();

        foreach (var hWnd in candidates)
            PreventWindowHiding(hWnd);

        var restored = Helper.RestoreHiddenExplorerWindows(removeLayeredStyle: !SettingsManager.HaveThemeIssue);
        if (restored > 0)
            DebugLog($"RecoverHiddenExplorerWindows reason={reason} restored={restored}");
    }

    private void PersistWindows()
    {
        var store = new List<WindowRecord>();
        lock (_closedWindowsLock)
        {
            if (SettingsManager.SaveClosedHistory) store.AddRange(_closedWindows);
            _closedWindows.Clear();
        }

        // Save currently open windows (explorer crash / system restart, logoff / AppExit)
        if (SettingsManager.RestorePreviousWindows)
            lock (_windowEntryDictLock)
            {
                store.AddRange(_windowEntryDict.Values
                    .Where(w => w.OnNavigateHandler != null)
                    .Select(w => new WindowRecord(w.Location!, name: w.Name!, restore: true)));
            }
        
        // DistinctBy location
        var distinctItems = store
            .GroupBy(w => w.Location)
            .Select(g => g.Last())
            .ToArray();
        
        // TakeLast 100
        SettingsManager.ClosedWindows = distinctItems.Skip(Math.Max(0, distinctItems.Length - 100)).ToArray();
    }

    public void Dispose()
    {
        DisposeShellObjects();
        _instanceRunning = false;
        _processWatcher.Dispose();
        GC.SuppressFinalize(this);
    }
}
