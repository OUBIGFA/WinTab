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
    private readonly ConcurrentDictionary<nint, byte> _registeredIndependentHWnds = new();
    private readonly ConcurrentDictionary<nint, TabStripBounds> _stripBoundsCache = new();
    private readonly ConcurrentDictionary<nint, byte> _stripBoundsRefreshInflight = new();
    private readonly ConcurrentDictionary<nint, (AutomationElement Element, long ExpiresAt)> _automationRootCache = new();
    private readonly DualKeyDictionary<InternetExplorer, nint?, WindowInfo> _windowEntryDict = [];
    private readonly List<WindowRecord> _closedWindows = new();
    private readonly object _windowEntryDictLock = new(), _closedWindowsLock = new(), _processLock = new();
    private readonly SemaphoreSlim _toOpenWindowsLock = new(1);
    private readonly SemaphoreSlim _pendingAutoMergeLock = new(1);
    private readonly ExplorerLaunchLocationResolver _locationResolver = new();
    private readonly ProcessWatcher _processWatcher;
    private int _mainExplorerProcessId;
    private Timer? _explorerCheckTimer;

    private WinEventHookThread? _winEventHookThread;
    private WinEventDelegate? _eventObjectCreateHookCallback;
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
        return TrySearchForTab(targetPath, out var tabHandle, out _) ? tabHandle : 0;
    }
    private bool TrySearchForTab(string targetPath, out nint tabHandle, out InternetExplorer? foundWindow)
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
                // Make sure it is not the newly created window
                if (!windowInfo.EventsHooked ||
                    !Helper.IsTimeUp(windowInfo.CreatedAt, 2_000) ||
                    !tab.HasValue ||
                    tab.Value == 0)
                {
                    continue;
                }

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

        return await SelectTabByUniqueNameVerified(windowHandle, tabHandle, Math.Min(timeoutMs, 350));
    }
    private async Task<bool> SelectTabByUniqueNameVerified(nint windowHandle, nint tabHandle, int timeoutMs, InternetExplorer? knownWindow = null)
    {
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

        if (!TrySelectSingleTabByAutomationName(windowHandle, tabName))
            return false;

        var activeTab = await Helper.DoUntilConditionAsync(
            () => GetActiveTabHandle(windowHandle),
            h => h == tabHandle,
            timeoutMs,
            10);

        return activeTab == tabHandle;
    }
    private static bool TrySelectSingleTabByAutomationName(nint windowHandle, string tabName)
    {
        try
        {
            var root = AutomationElement.FromHandle(windowHandle);
            if (root == null)
                return false;

            var tabItems = root.FindAll(
                    TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem))
                .Cast<AutomationElement>()
                .Select(element => new { Element = element, Rect = element.Current.BoundingRectangle, Name = element.Current.Name })
                .Where(item => !item.Rect.IsEmpty &&
                               item.Rect.Width >= 24 &&
                               item.Rect.Height >= 12 &&
                               IsMatchingTabName(item.Name, tabName))
                .OrderBy(item => item.Rect.Top)
                .ThenBy(item => item.Rect.Left)
                .Select(item => item.Element)
                .ToArray();

            return tabItems.Length == 1 && TrySelectAutomationElement(tabItems[0]);
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
        if (hWnd != 0)
            _registeredIndependentHWnds[hWnd] = 0;

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

        if (TryHideIncomingExplorerWindow(hWnd))
            SchedulePendingAutoMerge(50);
    }
    private void OnWindowCreated(nint hWinEventHook, uint eventType, nint hWnd, int idObject, int idChild, uint dwEventThread, uint dWmsEventTime)
    {
        if (!_isForcingTabs ||
            eventType != WinApi.EVENT_OBJECT_CREATE ||
            hWnd == 0 ||
            idObject != 0 ||
            idChild != 0)
        {
            return;
        }

        var rootWindow = WinApi.GetAncestor(hWnd, WinApi.GA_ROOT);
        if (rootWindow == 0)
            rootWindow = hWnd;

        if (TryHideIncomingExplorerWindow(rootWindow))
            SchedulePendingAutoMerge(50);
    }
    private bool TryHideIncomingExplorerWindow(nint hWnd)
    {
        if (!_isForcingTabs || hWnd == 0) return false;
        if (_processedHWnds.ContainsKey(hWnd) || _independentHWnds.ContainsKey(hWnd)) return false;
        if (Helper.IsCtrlShiftDown()) return false;
        if (!WinApi.IsWindowHasClassName(hWnd, "CabinetWClass")) return false;
        if (_mainWindowHandle != 0 && hWnd == _mainWindowHandle) return false;
        if (IsRegisteredIndependentExplorerWindow(hWnd)) return false;
        if (Helper.GetAnotherExplorerWindow(hWnd) == 0) return false;

        Helper.HideWindow(hWnd, SettingsManager.HaveThemeIssue);
        return true;
    }
    private bool IsRegisteredIndependentExplorerWindow(nint hWnd)
    {
        return _registeredIndependentHWnds.ContainsKey(hWnd);
    }
    private sealed class WinEventHookThread : IDisposable
    {
        private readonly WinEventDelegate _createCallback;
        private readonly WinEventDelegate _showCallback;
        private readonly ManualResetEventSlim _started = new();
        private readonly ManualResetEventSlim _stopped = new();
        private Thread? _thread;
        private uint _threadId;
        private nint _createHookId;
        private nint _showHookId;
        private bool _disposed;

        public WinEventHookThread(WinEventDelegate createCallback, WinEventDelegate showCallback)
        {
            _createCallback = createCallback;
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
                _createHookId = WinApi.SetWinEventHook(WinApi.EVENT_OBJECT_CREATE, WinApi.EVENT_OBJECT_CREATE, 0, _createCallback, 0, 0, hookFlags);
                _showHookId = WinApi.SetWinEventHook(WinApi.EVENT_OBJECT_SHOW, WinApi.EVENT_OBJECT_SHOW, 0, _showCallback, 0, 0, hookFlags);
                DebugLog($"WinEvent hook thread started id={_threadId} createHook={_createHookId} showHook={_showHookId}");
                _started.Set();

                while (WinApi.GetMessage(out var message, 0, 0, 0) > 0)
                {
                    WinApi.TranslateMessage(ref message);
                    WinApi.DispatchMessage(ref message);
                }
            }
            finally
            {
                if (_createHookId != 0)
                    WinApi.UnhookWinEvent(_createHookId);
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
    private void AdoptNewShellWindowsForImmediateConceal()
    {
        var shouldOpenAsWindow = Helper.IsCtrlShiftDown();
        var count = _shellWindows.Count;

        for (var i = count - 1; i >= 0; i--)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window) continue;

            WindowInfo windowInfo;
            nint hWnd;
            var canAutoMerge = false;
            var isKnownTopLevelTab = false;
            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.Keys.Contains(window)) continue;
                if (window.GetProperty("seenBefore") is not null) continue;

                hWnd = new IntPtr(window.HWND);
                if (hWnd != 0 && _windowEntryDict.Keys.Any(knownWindow =>
                    {
                        try
                        {
                            return new IntPtr(knownWindow.HWND) == hWnd;
                        }
                        catch
                        {
                            return false;
                        }
                    }))
                {
                    window.PutProperty("seenBefore", true);
                    isKnownTopLevelTab = true;
                    windowInfo = null!;
                }
                else
                {
                    window.PutProperty("seenBefore", true);
                    canAutoMerge = !shouldOpenAsWindow && _windowEntryDict.Count > 0;
                    windowInfo = new WindowInfo { CanAutoMerge = canAutoMerge };
                    _windowEntryDict.Add(window, windowInfo);

                    if (_windowEntryDict.Count == 1)
                    {
                        _mainWindowHandle = hWnd;

                        if (SettingsManager.RestorePreviousWindows && _closedWindows.Any(w => w.Restore))
                            _ = RestorePreviousWindows();
                    }
                }
            }

            if (isKnownTopLevelTab)
            {
                _ = TrackRegisteredTabWindow(window, hWnd);
                continue;
            }

            _ = GetTabHandle(window);

            if (!canAutoMerge)
            {
                _registeredIndependentHWnds[hWnd] = 0;
                if (shouldOpenAsWindow)
                    _independentHWnds[hWnd] = 0;

                PreventWindowHiding(hWnd);
                HookWindowEvents(window, windowInfo);
                continue;
            }

            Helper.HideWindow(hWnd, SettingsManager.HaveThemeIssue);
        }
    }
    private async Task TrackRegisteredTabWindow(InternetExplorer window, nint topLevelHWnd)
    {
        try
        {
            var tabHandle = await QueryTabHandle(window, updateDictionary: false);
            if (tabHandle == 0)
                return;

            WindowInfo windowInfo;
            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.Keys.Contains(window))
                    return;

                if (_windowEntryDict.TryGetValue(tabHandle, out InternetExplorer? _))
                    return;

                windowInfo = new WindowInfo();
                _windowEntryDict.Add(window, windowInfo, tabHandle);
            }

            HookWindowEvents(window, windowInfo);
            DebugLog($"Tracked registered tab hwnd={topLevelHWnd} tab={tabHandle}");
        }
        catch
        {
            //
        }
    }
    private void OnShellWindowRegistered(int __)
    {
        AdoptNewShellWindowsForImmediateConceal();
        SchedulePendingAutoMerge(50);
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
            var hadPendingCandidates = false;
            for (var i = 0; i < 32; i++)
            {
                AdoptNewShellWindowsForImmediateConceal();
                var candidates = GetPendingAutoMergeCandidates();
                if (candidates.Count == 0)
                    return false;

                hadPendingCandidates = true;
                var mergedAny = false;
                foreach (var (window, windowInfo) in candidates)
                {
                    if (await TryAutoMergePendingWindow(window, windowInfo))
                        mergedAny = true;
                }

                if (!mergedAny)
                    return true;
            }

            if (GetPendingAutoMergeCandidates().Count > 0)
                return hadPendingCandidates;

            return false;
        }
        finally
        {
            _pendingAutoMergeLock.Release();
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
        if (Helper.IsTimeUp(windowInfo.CreatedAt, 180_000))
        {
            TryShowAsIndependentWindow(window, windowInfo);
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
            _independentHWnds.ContainsKey(hWnd))
        {
            return false;
        }

        if (hWnd == _mainWindowHandle ||
            !Helper.IsFileExplorerWindow(hWnd) ||
            Helper.GetAnotherExplorerWindow(hWnd) == 0)
        {
            TryShowAsIndependentWindow(window, windowInfo);
            return false;
        }

        var tabCount = await WaitForExplorerTabCount(hWnd);
        if (tabCount != 1)
        {
            DebugLog($"Merge skip hwnd={hWnd} tabCount={tabCount}");
            return false;
        }

        var firstLocation = TryGetLocation(window);
        if (IsStartupExplorerLocation(firstLocation))
        {
            DebugLog($"Merge pending-startup hwnd={hWnd} location={firstLocation}");
            return false;
        }

        var location = await ResolveInitialLocation(window);
        DebugLog($"Merge resolved hwnd={hWnd} location={location}");
        if (string.IsNullOrWhiteSpace(location) ||
            location.StartsWith("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}"))
        {
            TryShowAsIndependentWindow(window, windowInfo);
            return false;
        }

        if (TryGetRecentlyClosedWindow(location, out var closedWindow))
        {
            windowInfo.CanAutoMerge = false;
            SelectItems(window, closedWindow!.SelectedItems);
            TryShowAsIndependentWindow(window, windowInfo);
            return false;
        }

        Helper.HideWindow(hWnd, SettingsManager.HaveThemeIssue);
        windowInfo.CanAutoMerge = false;

        var record = new WindowRecord(location, hWnd, GetSelectedItems(window));
        _ = OpenTabNavigateWithSelection(record, GetMainWindowHWnd(hWnd));

        DebugLog($"Merge dispatched hwnd={hWnd} location={location}");
        window.Quit();
        RemoveWindowAndUnhookEvents(window, windowInfo);
        return true;
    }
    private void TryShowAsIndependentWindow(InternetExplorer window, WindowInfo windowInfo)
    {
        nint hWnd;
        try
        {
            hWnd = new IntPtr(window.HWND);
        }
        catch
        {
            windowInfo.CanAutoMerge = false;
            return;
        }

        windowInfo.CanAutoMerge = false;
        PreventWindowHiding(hWnd);
        HookWindowEvents(window, windowInfo);
        _ = RestoreHiddenWindowAsync(hWnd);
    }
    private static async Task RestoreHiddenWindowAsync(nint hWnd)
    {
        await Helper.DoUntilNotDefaultAsync(() => Helper.ShowWindow(hWnd, removeCache: false), 1_500, 200);

        if (!SettingsManager.HaveThemeIssue)
            Helper.UpdateWindowLayered(hWnd, remove: true);

        // OnWindowShown can arrive after a release decision. Keep the cache briefly,
        // then allow later show events to be evaluated normally.
        _ = Task.Delay(3_000).ContinueWith(t => Helper.HiddenWindows.TryRemove(hWnd, out _), TaskScheduler.Default);
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
            _registeredIndependentHWnds.TryRemove(hWnd, out _);
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
        await _toOpenWindowsLock.WaitAsync();
        try
        {
            DebugLog($"OpenTab lock target={windowToOpen.Location}");
            if ((_reuseTabs || forceTabReuse) && !isDuplicate && _windowEntryDict.Count > 0 && !string.IsNullOrWhiteSpace(windowToOpen.Location))
            {
                if (TrySearchForTab(windowToOpen.Location, out var existingTab, out var existingWindow))
                {
                    windowHandle = WinApi.GetParent(existingTab);
                    if (await SelectTabByUniqueNameVerified(windowHandle, existingTab, 500, existingWindow))
                    {
                        WinApi.RestoreWindowToForeground(windowHandle);
                        DebugLog($"OpenTab reused target={windowToOpen.Location}");
                        return true;
                    }

                    DebugLog($"OpenTab reuse-select-failed target={windowToOpen.Location}");
                }
            }

            // Get the main window
            var mainWindowHWnd = Helper.IsFileExplorerWindow(windowHandle)
                ? windowHandle
                : GetMainWindowHWnd(windowToOpen.Handle);

            if (mainWindowHWnd == 0)
            {
                await OpenNewWindowWithSelection(windowToOpen, lockToOpenWindows: false);
                DebugLog($"OpenTab opened-window target={windowToOpen.Location}");
                return true;
            }

            for (var attempt = 1; attempt <= 3; attempt++)
            {
                nint newTabHandle = 0;
                try
                {
                    // Store the current tabs
                    var currentTabs = Helper.GetAllExplorerTabs(mainWindowHWnd).ToArray();
                    DebugLog($"OpenTab main={mainWindowHWnd} tabs={currentTabs.Length} attempt={attempt} target={windowToOpen.Location}");

                    // Request to open a new tab
                    await RequestToOpenNewTab(mainWindowHWnd, lockToOpenWindows: false);
                    DebugLog($"OpenTab requested main={mainWindowHWnd} attempt={attempt} target={windowToOpen.Location}");

                    // Wait for the new tab
                    newTabHandle = await Helper.ListenForNewExplorerTabAsync(mainWindowHWnd, currentTabs, 2_000);
                    if (newTabHandle == 0)
                    {
                        DebugLog($"OpenTab no-new-tab attempt={attempt} target={windowToOpen.Location}");
                        await Task.Delay(150);
                        continue;
                    }
                    DebugLog($"OpenTab new-tab={newTabHandle} attempt={attempt} target={windowToOpen.Location}");

                    var window = await Helper.DoUntilNotDefaultAsync(() => GetWindowByTabHandle(newTabHandle), 1_200, 50);
                    if (window == null)
                        window = await Helper.DoUntilNotDefaultAsync(() => FindShellWindowByTabHandle(newTabHandle, mainWindowHWnd), 3_500, 50);
                    if (window == null)
                    {
                        await CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle);
                        DebugLog($"OpenTab missing-shell-window tab={newTabHandle} attempt={attempt} target={windowToOpen.Location}");
                        continue;
                    }
                    DebugLog($"OpenTab found-shell-window tab={newTabHandle} attempt={attempt} target={windowToOpen.Location}");

                    if (!await NavigateNewTabToTargetAsync(window, windowToOpen.Location, attempt))
                    {
                        await CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle);
                        continue;
                    }

                    WinApi.RestoreWindowToForeground(mainWindowHWnd);
                    SelectItems(window, windowToOpen.SelectedItems);
                    return true;
                }
                catch (Exception ex)
                {
                    await CloseFailedNewTabAsync(mainWindowHWnd, newTabHandle);
                    DebugLog($"OpenTab attempt-error attempt={attempt} tab={newTabHandle} target={windowToOpen.Location} error={ex.GetType().Name}:{ex.Message}");
                }
            }

            return false;
        }
        finally
        {
            _toOpenWindowsLock.Release();
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
    private async Task<bool> NavigateNewTabToTargetAsync(InternetExplorer window, string targetLocation, int openAttempt)
    {
        for (var navigateAttempt = 1; navigateAttempt <= 3; navigateAttempt++)
        {
            if (!await NavigateToTargetIfNeeded(window, targetLocation))
            {
                DebugLog($"OpenTab navigate-failed openAttempt={openAttempt} navigateAttempt={navigateAttempt} target={targetLocation}");
                await Task.Delay(150);
                continue;
            }

            DebugLog($"OpenTab navigated openAttempt={openAttempt} navigateAttempt={navigateAttempt} target={targetLocation}");
            if (await WaitForNavigation(window, targetLocation, 5_000))
                return true;

            await Task.Delay(250);
            if (AreLocationsEquivalent(TryGetLocation(window), targetLocation))
                return true;

            DebugLog($"OpenTab target-check-failed openAttempt={openAttempt} navigateAttempt={navigateAttempt} target={targetLocation} current={TryGetLocation(window)}");
            await Task.Delay(150);
        }

        return false;
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
    private nint GetMainWindowHWnd(nint otherThan)
    {
        if (Helper.IsFileExplorerForeground(out var foregroundWindow) &&
            IsStableMergeTargetWindow(foregroundWindow, otherThan))
        {
            _mainWindowHandle = foregroundWindow;
            return _mainWindowHandle;
        }

        if (IsStableMergeTargetWindow(_mainWindowHandle, otherThan))
            return _mainWindowHandle;

        var allWindows = WinApi.FindAllWindowsEx("CabinetWClass");

        // Get another handle other than the newly created one. (In case if it is still alive.)
        _mainWindowHandle = allWindows
            .Where(h => IsStableMergeTargetWindow(h, otherThan))
            .Reverse() // To get the last one in the z-index (the oldest)
            .OrderByDescending(h => WinApi.FindAllWindowsEx("ShellTabWindowClass", h).Count()) // The one with the most tabs first
            .FirstOrDefault();

        if (_mainWindowHandle != 0) return _mainWindowHandle;

        _mainWindowHandle = allWindows
            .Where(h => IsFallbackMergeTargetWindow(h, otherThan))
            .Reverse()
            .OrderByDescending(h => WinApi.FindAllWindowsEx("ShellTabWindowClass", h).Count())
            .FirstOrDefault();

        return _mainWindowHandle;
    }
    private bool IsStableMergeTargetWindow(nint hWnd, nint otherThan)
    {
        if (hWnd == 0 || hWnd == otherThan)
            return false;
        if (!Helper.IsFileExplorerWindow(hWnd))
            return false;
        if (!IsRegisteredIndependentExplorerWindow(hWnd))
            return false;

        return Helper.GetAllExplorerTabs(hWnd).Take(1).Any();
    }
    private static bool IsFallbackMergeTargetWindow(nint hWnd, nint otherThan)
    {
        if (hWnd == 0 || hWnd == otherThan)
            return false;
        if (!Helper.IsFileExplorerWindow(hWnd))
            return false;

        return Helper.GetAllExplorerTabs(hWnd).Take(1).Any();
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

        if (Helper.IsFileExplorerForeground(out var foregroundWindow))
            _mainWindowHandle = foregroundWindow;
        
        if (SettingsManager.ClosedWindows != null)
            lock (_closedWindowsLock) _closedWindows.AddRange(SettingsManager.ClosedWindows);

        // Hook the global "WindowRegistered" event
        _windowRegisteredHandler = OnShellWindowRegistered;
        _shellWindows.WindowRegistered += _windowRegisteredHandler;

        // Match ExplorerTabUtility's event-driven model: WinEvent hides the new
        // Explorer window, and ShellWindows.WindowRegistered performs the merge.
        _eventObjectCreateHookCallback = OnWindowCreated;
        _eventObjectShowHookCallback = OnWindowShown;
        _winEventHookThread = new WinEventHookThread(_eventObjectCreateHookCallback, _eventObjectShowHookCallback);
        _winEventHookThread.Start();

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
            _registeredIndependentHWnds[new IntPtr(window.HWND)] = 0;

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
        _eventObjectCreateHookCallback = null;
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
