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
    private const int StripBoundsRectMatchSlop = 2;
    private static readonly string? DebugLogPath = Environment.GetEnvironmentVariable("WINTAB_DEBUG_LOG");
    private static bool _instanceRunning;
    private static Guid _shellBrowserGuid = typeof(IShellBrowser).GUID;

    private ShellWindows _shellWindows = null!;
    private ShellPathComparer _shellPathComparer = null!;
    private StaTaskScheduler _staTaskScheduler = null!;
    private nint _mainWindowHandle;
    private readonly ConcurrentDictionary<nint, byte> _processedHWnds = new();
    private readonly ConcurrentDictionary<nint, byte> _independentHWnds = new();
    private readonly ConcurrentDictionary<nint, TabStripBounds> _stripBoundsCache = new();
    private readonly ConcurrentDictionary<nint, byte> _stripBoundsRefreshInflight = new();
    private readonly ConcurrentDictionary<nint, (AutomationElement Element, long ExpiresAt)> _automationRootCache = new();
    private readonly DualKeyDictionary<InternetExplorer, nint?, WindowInfo> _windowEntryDict = [];
    private readonly List<WindowRecord> _closedWindows = new();
    private readonly object _windowEntryDictLock = new(), _closedWindowsLock = new(), _processLock = new();
    private readonly SemaphoreSlim _toOpenWindowsLock = new(1);
    private readonly SemaphoreSlim _windowRegistrationLock = new(1);
    private readonly SemaphoreSlim _pendingAutoMergeLock = new(1);
    private readonly ExplorerLaunchLocationResolver _locationResolver = new();
    private readonly ProcessWatcher _processWatcher;
    private int _mainExplorerProcessId;
    private Timer? _explorerCheckTimer;

    private nint _eventObjectShowHookId;
    private WinEventDelegate? _eventObjectShowHookCallback;
    private DShellWindowsEvents_WindowRegisteredEventHandler? _windowRegisteredHandler;

    private string _defaultLocation = null!;
    private bool _reuseTabs = true;
    private bool _isForcingTabs;
    private int _pendingAutoMergeScheduled;
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
        _isForcingTabs = true;
        DebugLog("StartHook");
    }

    public void StopHook()
    {
        if (!_isForcingTabs) return;
        _isForcingTabs = false;
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
        nint targetPidl = 0;
        try
        {
            targetPidl = _shellPathComparer.GetPidlFromPath(targetPath);
            if (targetPidl == 0) return 0;

            foreach (var (window, windowInfo, tabHandle) in _windowEntryDict)
            {
                // Make sure it is not the newly created window
                if (!Helper.IsTimeUp(windowInfo.CreatedAt, 2_000) || !tabHandle.HasValue || tabHandle.Value == 0)
                    continue;

                var comparePath = windowInfo.Location ?? GetLocation(window);

                if (_shellPathComparer.IsEquivalent(targetPath, comparePath, targetPidl))
                    return tabHandle.Value;
            }

            return 0;
        }
        catch
        {
            return 0;
        }
        finally
        {
            if (targetPidl != 0)
                Marshal.FreeCoTaskMem(targetPidl);
        }
    }
    public async Task SelectTabByHandle(nint windowHandle, nint tabHandle)
    {
        var tabs = Helper.GetAllExplorerTabs(windowHandle).ToArray();
        if (tabs.Length == 0) return;

        var activeTab = tabs[0];
        for (var i = 0; i < tabs.Length; i++)
        {
            if (activeTab == tabHandle) break;

            SelectTabByIndex(windowHandle, i);

            // ReSharper disable once AccessToModifiedClosure
            activeTab = await Helper.DoUntilConditionAsync(
                () => WinApi.FindWindowEx(windowHandle, 0, "ShellTabWindowClass", null),
                h => h != activeTab);
        }
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
        if (!_isForcingTabs || idObject != 0 || idChild != 0) return;
        
        // Check if the hWnd was processed by OnShellWindowRegistered
        if (_processedHWnds.TryRemove(hWnd, out _)) return;
        
        if (!WinApi.IsWindowHasClassName(hWnd, "CabinetWClass")) return;

        if (Helper.IsCtrlShiftDown() || Helper.GetAnotherExplorerWindow(hWnd) == 0) return;
        Helper.HideWindow(hWnd, SettingsManager.HaveThemeIssue);
        SchedulePendingAutoMerge();
    }
    private InternetExplorer? GetRecentlyCreatedWindow(out WindowInfo? windowInfo)
    {
        // When a new window is registered, it's typically the last in the collection
        var count = _shellWindows.Count;
        for (var i = count - 1; i >= 0; i--)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window) continue;

            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.Keys.Contains(window)) continue;
                if (window.GetProperty("seenBefore") is not null) continue;
                window.PutProperty("seenBefore", true);

                windowInfo = new WindowInfo { CanAutoMerge = true };
                _windowEntryDict.Add(window, windowInfo);

                if (_windowEntryDict.Count == 1)
                {
                    _mainWindowHandle = new IntPtr(window.HWND);

                    if (SettingsManager.RestorePreviousWindows && _closedWindows.Any(w => w.Restore))
                        _ = RestorePreviousWindows();
                }
                
                return window;
            }
        }

        windowInfo = null;
        return null;
    }
    private async void OnShellWindowRegistered(int __)
    {
        await _windowRegistrationLock.WaitAsync();
        try
        {
            await ProcessShellWindowRegisteredAsync();
        }
        finally
        {
            _windowRegistrationLock.Release();
        }

        SchedulePendingAutoMerge();
    }
    private async Task ProcessShellWindowRegisteredAsync()
    {
        var showAgain = true;
        nint hWnd = 0;
        try
        {
            var shouldOpenAsWindow = Helper.IsCtrlShiftDown();

            WindowInfo windowInfo = null!;
            var window = await Helper.DoUntilNotDefaultAsync(() => GetRecentlyCreatedWindow(out windowInfo!), 2_500, 70);
            if (window == null) return;

            _ = GetTabHandle(window);

            hWnd = new IntPtr(window.HWND);
            DebugLog($"Registered hwnd={hWnd}");
            
            if (shouldOpenAsWindow)
            {
                windowInfo.CanAutoMerge = false;
                _independentHWnds[hWnd] = 0;
                PreventWindowHiding(hWnd);
                HookWindowEvents(window, windowInfo);
                return;
            }
            
            var location = await ResolveInitialLocation(window);
            DebugLog($"Resolved hwnd={hWnd} location={location}");

            //Control Panel
            if (location.StartsWith("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}"))
            {
                windowInfo.CanAutoMerge = false;
                _independentHWnds[hWnd] = 0;
                PreventWindowHiding(hWnd);
                RemoveWindowAndUnhookEvents(window, windowInfo);
                return;
            }

            var tabCount = await WaitForExplorerTabCount(hWnd);
            DebugLog($"Decision hwnd={hWnd} tabCount={tabCount} dict={_windowEntryDict.Count} main={_mainWindowHandle}");

            // Check if this is a single tab window and there are other windows
            var shouldReopenAsTab = (_isForcingTabs || _reuseTabs) &&
                                    _windowEntryDict.Count > 1 &&
                                    hWnd != _mainWindowHandle &&
                                    tabCount == 1;

            if (shouldReopenAsTab)
                Helper.HideWindow(hWnd, SettingsManager.HaveThemeIssue);
            else
                PreventWindowHiding(hWnd);

            // Check if it is a detached tab
            var isRecentlyClosed = TryGetRecentlyClosedWindow(location, out var closedWindow);
            if (isRecentlyClosed)
            {
                windowInfo.CanAutoMerge = false;
                SelectItems(window, closedWindow!.SelectedItems);
            }

            shouldReopenAsTab = shouldReopenAsTab && !isRecentlyClosed;

            if (shouldReopenAsTab)
            {
                var reopenedAsTab = await OpenTabNavigateWithSelection(
                    new WindowRecord(location, hWnd, GetSelectedItems(window)),
                    _mainWindowHandle);

                if (reopenedAsTab)
                {
                    DebugLog($"Merged registered hwnd={hWnd} location={location}");
                    showAgain = false;

                    window.Quit();
                    RemoveWindowAndUnhookEvents(window, windowInfo);
                    return;
                }

                DebugLog($"Merge failed registered hwnd={hWnd} location={location}");
            }

            // OnQuit might fire after ShellWindowRegistered in case of reattached tab (and there were selected files)
            if (!isRecentlyClosed)
            {
                isRecentlyClosed = await Helper.DoUntilNotDefaultAsync(() => TryGetRecentlyClosedWindow(location, out closedWindow), 700, 50);
                if (isRecentlyClosed)
                {
                    windowInfo.CanAutoMerge = false;
                    SelectItems(window, closedWindow!.SelectedItems);
                }
            }

            HookWindowEvents(window, windowInfo);
        }
        catch {/**/}
        finally
        {
            if (showAgain)
            {
                await Helper.DoUntilNotDefaultAsync(() => Helper.ShowWindow(hWnd, removeCache: false), 1_500, 200);

                if (!SettingsManager.HaveThemeIssue)
                    Helper.UpdateWindowLayered(hWnd, remove: true);

                // OnWindowShown might fire after ShellWindowRegistered and hide it again, keep the cache, wait a bit, then remove it.
                _ = Task.Delay(3000).ContinueWith(t => Helper.HiddenWindows.TryRemove(hWnd, out _), TaskScheduler.Default);
            }
        }
    }
    private Task<string> ResolveInitialLocation(InternetExplorer window)
    {
        return _locationResolver.ResolveAsync(
            () => TryGetLocation(window),
            IsStartupExplorerLocation);
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
    private bool IsStartupExplorerLocation(string location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return true;

        if (StringComparer.OrdinalIgnoreCase.Equals(location, _defaultLocation))
            return true;

        try
        {
            return _shellPathComparer.IsEquivalent(location, _defaultLocation);
        }
        catch
        {
            return false;
        }
    }
    private static Task<int> WaitForExplorerTabCount(nint hWnd)
    {
        return Helper.DoUntilConditionAsync(
            () => Helper.GetAllExplorerTabs(hWnd).Take(2).Count(),
            count => count > 0,
            1_500,
            25);
    }
    private void SchedulePendingAutoMerge(int delayMs = 350)
    {
        if (!_isForcingTabs && !_reuseTabs)
            return;

        if (Interlocked.Exchange(ref _pendingAutoMergeScheduled, 1) == 1)
            return;

        _ = Task.Run(async () =>
        {
            var shouldRetry = false;
            try
            {
                await Task.Delay(delayMs);
                shouldRetry = await MergePendingAutoMergeWindowsAsync();
            }
            catch
            {
                //
            }
            finally
            {
                Interlocked.Exchange(ref _pendingAutoMergeScheduled, 0);
            }

            if (shouldRetry)
                SchedulePendingAutoMerge();
        });
    }
    private async Task<bool> MergePendingAutoMergeWindowsAsync()
    {
        if (_shellWindows == null || (!_isForcingTabs && !_reuseTabs))
            return false;

        if (!await _pendingAutoMergeLock.WaitAsync(0))
            return true;

        try
        {
            if (!await _windowRegistrationLock.WaitAsync(0))
                return true;

            try
            {
                var hadPendingCandidates = false;
                for (var i = 0; i < 8; i++)
                {
                    AdoptUntrackedShellWindowsForAutoMerge();
                    var candidates = GetPendingAutoMergeCandidates();
                    if (candidates.Count == 0)
                        return false;

                    hadPendingCandidates = true;
                    var mergedAny = false;
                    foreach (var (window, windowInfo) in candidates)
                    {
                        if (!await TryAutoMergePendingWindow(window, windowInfo))
                            continue;

                        mergedAny = true;
                        break;
                    }

                    if (!mergedAny)
                        return hadPendingCandidates;
                }

                return hadPendingCandidates;
            }
            finally
            {
                _windowRegistrationLock.Release();
            }
        }
        finally
        {
            _pendingAutoMergeLock.Release();
        }
    }
    private void AdoptUntrackedShellWindowsForAutoMerge()
    {
        if (_shellWindows == null)
            return;

        var count = _shellWindows.Count;
        for (var i = count - 1; i >= 0; i--)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;

            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.Keys.Contains(window))
                    continue;

                if (window.GetProperty("seenBefore") is not null)
                    continue;

                nint hWnd;
                try
                {
                    hWnd = new IntPtr(window.HWND);
                }
                catch
                {
                    continue;
                }

                if (!Helper.IsFileExplorerWindow(hWnd))
                    continue;

                window.PutProperty("seenBefore", true);

                var windowInfo = new WindowInfo { CanAutoMerge = true };
                _windowEntryDict.Add(window, windowInfo);
                DebugLog($"Adopted hwnd={hWnd}");
                _ = GetTabHandle(window);
            }
        }
    }
    private List<(InternetExplorer Window, WindowInfo WindowInfo)> GetPendingAutoMergeCandidates()
    {
        lock (_windowEntryDictLock)
        {
            var candidates = new List<(InternetExplorer Window, WindowInfo WindowInfo)>();
            foreach (var (window, windowInfo, _) in _windowEntryDict)
            {
                if (windowInfo.CanAutoMerge)
                    candidates.Add((window, windowInfo));
            }

            return candidates;
        }
    }
    private async Task<bool> TryAutoMergePendingWindow(InternetExplorer window, WindowInfo windowInfo)
    {
        windowInfo.AutoMergeAttempts++;
        if (windowInfo.AutoMergeAttempts > 20)
        {
            windowInfo.CanAutoMerge = false;
            HookWindowEvents(window, windowInfo);
            return false;
        }

        nint hWnd;
        try
        {
            hWnd = new IntPtr(window.HWND);
        }
        catch
        {
            return false;
        }

        if (!windowInfo.CanAutoMerge ||
            _independentHWnds.ContainsKey(hWnd) ||
            hWnd == _mainWindowHandle ||
            !Helper.IsFileExplorerWindow(hWnd) ||
            Helper.GetAnotherExplorerWindow(hWnd) == 0)
        {
            return false;
        }

        var tabCount = await WaitForExplorerTabCount(hWnd);
        if (tabCount != 1)
        {
            DebugLog($"Sweep skip hwnd={hWnd} tabCount={tabCount}");
            return false;
        }

        var location = await ResolveInitialLocation(window);
        DebugLog($"Sweep resolved hwnd={hWnd} location={location}");
        if (string.IsNullOrWhiteSpace(location) ||
            location.StartsWith("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}"))
        {
            windowInfo.CanAutoMerge = false;
            HookWindowEvents(window, windowInfo);
            return false;
        }

        if (TryGetRecentlyClosedWindow(location, out var closedWindow))
        {
            windowInfo.CanAutoMerge = false;
            SelectItems(window, closedWindow!.SelectedItems);
            HookWindowEvents(window, windowInfo);
            return false;
        }

        Helper.HideWindow(hWnd, SettingsManager.HaveThemeIssue);

        var reopenedAsTab = await OpenTabNavigateWithSelection(
            new WindowRecord(location, hWnd, GetSelectedItems(window)),
            _mainWindowHandle);

        if (!reopenedAsTab)
        {
            DebugLog($"Sweep merge failed hwnd={hWnd} location={location}");
            return false;
        }

        DebugLog($"Sweep merged hwnd={hWnd} location={location}");
        window.Quit();
        RemoveWindowAndUnhookEvents(window, windowInfo);
        return true;
    }
    private void HookWindowEvents(InternetExplorer window, WindowInfo windowInfo)
    {
        if (windowInfo.EventsHooked)
            return;

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
            lock (_windowEntryDictLock)
                _windowEntryDict.Remove(window);
        }
    }
    private void RemoveWindowAndUnhookEvents(InternetExplorer window, WindowInfo windowInfo, bool useLock = true)
    {
        // Unsubscribe
        if (windowInfo.OnQuitHandler != null) window.OnQuit -= windowInfo.OnQuitHandler;
        if (windowInfo.OnNavigateHandler != null) window.NavigateComplete2 -= windowInfo.OnNavigateHandler;

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
            _processedHWnds.TryRemove(hWnd, out _);
            _independentHWnds.TryRemove(hWnd, out _);
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
    private async Task<bool> OpenTabNavigateWithSelection(WindowRecord windowToOpen, nint windowHandle = 0, bool isDuplicate = false, bool forceTabReuse = false)
    {
        await _toOpenWindowsLock.WaitAsync();
        try
        {
            if ((_reuseTabs || forceTabReuse) && !isDuplicate && _windowEntryDict.Count > 0 && !string.IsNullOrWhiteSpace(windowToOpen.Location))
            {
                var existingTab = SearchForTab(windowToOpen.Location);
                if (existingTab != 0)
                {
                    windowHandle = WinApi.GetParent(existingTab);
                    await SelectTabByHandle(windowHandle, existingTab);
                    WinApi.RestoreWindowToForeground(windowHandle);
                    return true;
                }
            }

            // Get the main window
            var mainWindowHWnd = Helper.IsFileExplorerWindow(windowHandle)
                ? windowHandle
                : GetMainWindowHWnd(windowToOpen.Handle);

            if (mainWindowHWnd == 0)
            {
                await OpenNewWindowWithSelection(windowToOpen, lockToOpenWindows: false);
                return true;
            }

            // Store the current tabs
            var currentTabs = Helper.GetAllExplorerTabs(mainWindowHWnd).ToArray();

            // Request to open a new tab
            await RequestToOpenNewTab(mainWindowHWnd, lockToOpenWindows: false);

            // Wait for the new tab
            var newTabHandle = await Helper.ListenForNewExplorerTabAsync(mainWindowHWnd, currentTabs, 2_000);
            if (newTabHandle == 0) return false;

            // Get the window object. Do not rely on WindowRegistered here: this method can be
            // called while the registration pipeline is intentionally serialized.
            var window = await Helper.DoUntilNotDefaultAsync(() => FindShellWindowByTabHandle(newTabHandle), 3_000, 50);
            if (window == null) return false;

            try
            {
                await Navigate(window, windowToOpen.Location);
            }
            catch
            {
                return false;
            }

            WinApi.RestoreWindowToForeground(mainWindowHWnd);

            var reachedTarget = await WaitForNavigation(window, windowToOpen.Location);
            if (!reachedTarget)
                return false;

            SelectItems(window, windowToOpen.SelectedItems);
            return true;
        }
        finally
        {
            _toOpenWindowsLock.Release();
        }
    }
    private async Task<bool> WaitForNavigation(InternetExplorer window, string targetLocation)
    {
        if (string.IsNullOrWhiteSpace(targetLocation))
            return true;

        var resolvedLocation = await Helper.DoUntilConditionAsync(
            () => TryGetLocation(window),
            location => AreLocationsEquivalent(location, targetLocation),
            5_000,
            50);

        return AreLocationsEquivalent(resolvedLocation, targetLocation);
    }
    private bool AreLocationsEquivalent(string location, string targetLocation)
    {
        if (string.IsNullOrWhiteSpace(location) || string.IsNullOrWhiteSpace(targetLocation))
            return false;

        if (StringComparer.OrdinalIgnoreCase.Equals(location, targetLocation))
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
    private nint GetMainWindowHWnd(nint otherThan)
    {
        if (Helper.IsFileExplorerWindow(_mainWindowHandle))
            return _mainWindowHandle;

        var allWindows = WinApi.FindAllWindowsEx("CabinetWClass");

        // Get another handle other than the newly created one. (In case if it is still alive.)
        _mainWindowHandle = allWindows
            .Where(h => h != otherThan)
            .Reverse() // To get the last one in the z-index (the oldest)
            .OrderByDescending(h => WinApi.FindAllWindowsEx("ShellTabWindowClass", h).Count()) // The one with the most tabs first
            .FirstOrDefault();

        if (_mainWindowHandle != 0) return _mainWindowHandle;

        return Helper.IsFileExplorerWindow(otherThan) ? otherThan : 0;
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
    private async Task<InternetExplorer?> FindShellWindowByTabHandle(nint tabHandle)
    {
        var count = _shellWindows.Count;
        for (var i = count - 1; i >= 0; i--)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;

            var currentTabHandle = await QueryTabHandle(window, updateDictionary: false);
            if (currentTabHandle != tabHandle)
                continue;

            WindowInfo windowInfo;
            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.TryGetValue(window, out windowInfo!))
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

            HookWindowEvents(window, windowInfo);
            return window;
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
        
        if (SettingsManager.ClosedWindows != null)
            lock (_closedWindowsLock) _closedWindows.AddRange(SettingsManager.ClosedWindows);

        // Hook the global "WindowRegistered" event
        _windowRegisteredHandler = OnShellWindowRegistered;
        _shellWindows.WindowRegistered += _windowRegisteredHandler;

        // Hook the global "OBJECT_SHOW" event
        _eventObjectShowHookCallback = OnWindowShown;
        _eventObjectShowHookId = WinApi.SetWinEventHook(WinApi.EVENT_OBJECT_SHOW, WinApi.EVENT_OBJECT_SHOW, 0, _eventObjectShowHookCallback, 0, 0, 0);

        // Hook the event handlers for already-open windows
        var hasOpen = false;
        var count = _shellWindows.Count;
        for (var i = 0; i < count; i++)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;
            hasOpen = true;

            var windowInfo = new WindowInfo();
            _windowEntryDict.Add(window, windowInfo);
            window.PutProperty("seenBefore", true);

            _ = GetTabHandle(window);
            HookWindowEvents(window, windowInfo);
        }

        if (!hasOpen) return;
        lock (_closedWindowsLock)
            foreach (var window in _closedWindows) window.Restore = false;
    }
    private void DisposeShellObjects()
    {
        if (_shellWindows == null)
            return;

        PersistWindows();

        // Unhook global event
        if (_windowRegisteredHandler != null)
        {
            _shellWindows.WindowRegistered -= _windowRegisteredHandler;
            _windowRegisteredHandler = null;
        }
        if (_eventObjectShowHookCallback != null)
        {
            WinApi.UnhookWinEvent(_eventObjectShowHookId);
            _eventObjectShowHookCallback = null;
        }

        // Unsubscribe from each InternetExplorer instance's events
        foreach (var (window, windowInfo) in _windowEntryDict)
        {
            // Unsubscribe
            if (windowInfo.OnQuitHandler != null) window.OnQuit -= windowInfo.OnQuitHandler;
            if (windowInfo.OnNavigateHandler != null) window.NavigateComplete2 -= windowInfo.OnNavigateHandler;

            // Release the COM object
            Marshal.ReleaseComObject(window);
        }
        _windowEntryDict.Clear();

        // Release the ShellWindows COM object
        Marshal.ReleaseComObject(_shellWindows);

        _shellPathComparer.Dispose();
        _staTaskScheduler.Dispose();
        _shellWindows = null!;
        _shellPathComparer = null!;
        _staTaskScheduler = null!;
        _mainWindowHandle = 0;
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
