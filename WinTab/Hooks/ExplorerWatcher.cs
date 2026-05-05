using SHDocVw;
using Shell32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Windows.Automation;
using WinTab.Helpers;
using WinTab.Interop;
using WinTab.Managers;
using WinTab.Models;
using WinTab.WinAPI;

namespace WinTab.Hooks;

using WindowEntry = DualKeyEntry<InternetExplorer, nint?, WindowInfo>;

public sealed class ExplorerWatcher : IHook
{
    private const int MergeHideGraceMs = 2_500;
    private const int AutomationRootTtlMs = 5_000;
    private const int StripBoundsRectMatchSlop = 2; // px tolerance when comparing window rect snapshots
    private static bool _instanceRunning;
    private static Guid ShellBrowserGuid = typeof(IShellBrowser).GUID;

    private readonly ProcessWatcher _processWatcher;
    private readonly ConcurrentDictionary<nint, byte> _processedHWnds = new();
    private readonly ConcurrentDictionary<nint, byte> _knownExplorerWindows = new();
    private readonly ConcurrentDictionary<nint, long> _mergeHiddenWindows = new();
    private readonly ConcurrentDictionary<nint, TabStripBounds> _stripBoundsCache = new();
    private readonly ConcurrentDictionary<nint, byte> _stripBoundsRefreshInflight = new();
    private readonly ConcurrentDictionary<nint, (AutomationElement element, long expiresAt)> _automationRootCache = new();
    private readonly DualKeyDictionary<InternetExplorer, nint?, WindowInfo> _windowEntryDict = [];
    private readonly object _windowEntryDictLock = new();
    private readonly object _processLock = new();
    private readonly SemaphoreSlim _toOpenWindowsLock = new(1, 1);

    private ShellWindows? _shellWindows;
    private ShellPathComparer? _shellPathComparer;
    private StaTaskScheduler? _staTaskScheduler;
    private Timer? _explorerCheckTimer;
    private Timer? _hiddenRecoveryTimer;
    private DShellWindowsEvents_WindowRegisteredEventHandler? _windowRegisteredHandler;
    private int _mainExplorerProcessId;
    private nint _mainWindowHandle;
    private string _defaultLocation = string.Empty;
    private bool _reuseTabs = true;
    private bool _isForcingTabs;
    private long _suppressPreHideUntil;

    private sealed record TabStripBounds(Rectangle StripRect, Rectangle[] TabRects, RECT WindowRect, long RefreshedAt);

    public ExplorerWatcher()
    {
        if (_instanceRunning)
            throw new InvalidOperationException("Only one instance of ExplorerWatcher is allowed at a time.");

        _instanceRunning = true;
        _processWatcher = new ProcessWatcher("explorer");
        _processWatcher.ProcessTerminated += OnExplorerProcessTerminated;
        StartExplorerProcessCheck();
    }

    public bool IsHookActive => _isForcingTabs;
    public bool IsShellReady => _mainExplorerProcessId != 0 && _shellWindows != null;
    public event Action? OnShellInitialized;
    public event Action<string>? StatusChanged;

    public void StartHook() => _isForcingTabs = true;
    public void StopHook() => _isForcingTabs = false;
    public void SetReuseTabs(bool reuseTabs) => _reuseTabs = reuseTabs;

    public bool TryCloseTabAtScreenPoint(Point screenPoint, nint explorerWindow = 0)
    {
        explorerWindow = Helper.IsFileExplorerWindow(explorerWindow) ? explorerWindow : WinApi.GetForegroundWindow();
        if (!Helper.IsFileExplorerWindow(explorerWindow))
            return false;

        var tabItem = GetExplorerTabItemAtPoint(screenPoint, explorerWindow);
        if (tabItem == null)
            return false;

        SuppressPreHideForTabClose();
        if (!InvokeTabCloseButton(tabItem))
            return false;

        _ = RestoreExplorerAfterNativeCloseAsync(explorerWindow);
        return true;
    }

    public bool IsExplorerTabTitleAtScreenPoint(Point screenPoint, nint explorerWindow = 0)
    {
        explorerWindow = Helper.IsFileExplorerWindow(explorerWindow) ? explorerWindow : WinApi.GetForegroundWindow();
        return Helper.IsFileExplorerWindow(explorerWindow) && GetExplorerTabItemAtPoint(screenPoint, explorerWindow) != null;
    }

    /// <summary>
    /// Cheap, hook-thread-safe predicate. Returns true ONLY when <paramref name="screenPoint"/> sits
    /// inside an actual Explorer tab title (TabItem rect) of <paramref name="explorerWindow"/>. The
    /// "+" new-tab button, the empty drag area to the right of the last tab, the title bar, and the
    /// minimize/maximize/close buttons are all explicitly excluded so a double-click outside a real
    /// tab keeps Explorer's native behaviour (maximize/restore on title bar, drag, etc.).
    /// Tab rects come from a per-window UI Automation snapshot computed on a background thread; the
    /// cache is warmed eagerly when the window is registered. If the cache isn't ready yet we return
    /// false (no heuristic fallback) — losing one tab-close gesture is far better than swallowing a
    /// title-bar maximize.
    /// </summary>
    public bool IsPointOnTabStrip(Point screenPoint, nint explorerWindow)
    {
        if (!Helper.IsFileExplorerWindow(explorerWindow))
            return false;

        if (!WinApi.GetWindowRect(explorerWindow, out var winRect))
            return false;

        if (screenPoint.X < winRect.Left || screenPoint.X >= winRect.Right ||
            screenPoint.Y < winRect.Top)
            return false;

        if (_stripBoundsCache.TryGetValue(explorerWindow, out var bounds) &&
            RectsApproxEqual(bounds.WindowRect, winRect))
        {
            // Refresh in the background if the snapshot is getting stale (tabs may have been added,
            // removed, or reordered without resizing the window). The inflight guard makes this cheap.
            if (Environment.TickCount64 - bounds.RefreshedAt > 1_500)
                ScheduleStripBoundsRefresh(explorerWindow);

            foreach (var tabRect in bounds.TabRects)
            {
                if (tabRect.Contains(screenPoint.X, screenPoint.Y))
                    return true;
            }

            return false;
        }

        if (bounds != null)
            _stripBoundsCache.TryRemove(explorerWindow, out _);

        ScheduleStripBoundsRefresh(explorerWindow);
        return false;
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

        var tabRects = new List<Rectangle>(tabItems.Count);
        AutomationElement? firstTabItem = null;
        foreach (AutomationElement tab in tabItems)
        {
            var rect = tab.Current.BoundingRectangle;
            if (rect.IsEmpty || rect.Width < 24 || rect.Height < 12)
                continue;

            firstTabItem ??= tab;
            tabRects.Add(new Rectangle((int)rect.X, (int)rect.Y, (int)rect.Width, (int)rect.Height));
        }

        if (firstTabItem == null || tabRects.Count == 0)
            return;

        // Walk up to the Tab control whose BoundingRectangle covers the entire strip — used only as a
        // diagnostic / cache-key for window-rect mismatch detection. Hit-testing uses tabRects only.
        var walker = TreeWalker.ControlViewWalker;
        var element = firstTabItem;
        System.Windows.Rect? stripRect = null;
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
        {
            var firstRect = firstTabItem.Current.BoundingRectangle;
            if (firstRect.IsEmpty)
                return;
            stripRect = firstRect;
        }

        if (!WinApi.GetWindowRect(explorerWindow, out var winSnapshot))
            return;

        var s = stripRect.Value;
        _stripBoundsCache[explorerWindow] = new TabStripBounds(
            new Rectangle((int)s.X, (int)s.Y, (int)s.Width, (int)s.Height),
            tabRects.ToArray(),
            winSnapshot,
            Environment.TickCount64);
    }

    private void InvalidateStripBounds(nint hWnd)
    {
        _stripBoundsCache.TryRemove(hWnd, out _);
        _stripBoundsRefreshInflight.TryRemove(hWnd, out _);
    }

    /// <summary>
    /// Force-refresh the cached tab strip bounds for an Explorer window. Called right after a
    /// programmatic tab close so the next double-click sees the post-close tab layout (tabs slide
    /// over to fill the gap).
    /// </summary>
    public void RefreshTabStripBounds(nint explorerWindow)
    {
        InvalidateStripBounds(explorerWindow);
        if (Helper.IsFileExplorerWindow(explorerWindow))
            ScheduleStripBoundsRefresh(explorerWindow);
    }

    private AutomationElement? GetAutomationRoot(nint explorerWindow)
    {
        var now = Environment.TickCount64;
        if (_automationRootCache.TryGetValue(explorerWindow, out var entry) && entry.expiresAt > now)
            return entry.element;

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

    private void InvalidateAutomationRoot(nint hWnd)
    {
        _automationRootCache.TryRemove(hWnd, out _);
    }

    private static async Task RestoreExplorerAfterNativeCloseAsync(nint explorerWindow)
    {
        await Task.Delay(500);
        RestoreExplorerAfterTabClose(explorerWindow);
    }

    private static bool InvokeTabCloseButton(AutomationElement tabItem)
    {
        try
        {
            var closeButton = tabItem.FindFirst(
                TreeScope.Descendants,
                new OrCondition(
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "CloseButton"),
                    new PropertyCondition(AutomationElement.NameProperty, "Close tab"),
                    new PropertyCondition(AutomationElement.NameProperty, "\u5173\u95ed\u6807\u7b7e\u9875")));

            if (closeButton == null)
                return false;

            if (closeButton.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern) &&
                pattern is InvokePattern invokePattern)
            {
                invokePattern.Invoke();
                return true;
            }
        }
        catch
        {
            //
        }

        return false;
    }

    private void SuppressPreHideForTabClose()
    {
        Interlocked.Exchange(ref _suppressPreHideUntil, Environment.TickCount64 + 5_000);
    }

    private static void RestoreExplorerAfterTabClose(nint explorerWindow)
    {
        if (!Helper.IsFileExplorerWindow(explorerWindow))
            return;

        if (Helper.HiddenWindows.ContainsKey(explorerWindow))
            Helper.ShowWindow(explorerWindow, removeCache: true);
        else
            WinApi.ShowWindow(explorerWindow, WinApi.SW_SHOWNOACTIVATE);

        WinApi.RestoreWindowToForeground(explorerWindow);
    }

    public nint SearchForTab(string targetPath)
    {
        if (_shellPathComparer == null || string.IsNullOrWhiteSpace(targetPath))
            return 0;

        nint targetPidl = 0;
        try
        {
            targetPidl = _shellPathComparer.GetPidlFromPath(targetPath);
            if (targetPidl == 0)
                return 0;

            foreach (var (window, windowInfo, tabHandle) in _windowEntryDict)
            {
                if (!Helper.IsTimeUp(windowInfo.CreatedAt, 2_000) || !tabHandle.HasValue || tabHandle.Value == 0)
                    continue;

                var comparePath = windowInfo.Location ?? TryGetLocation(window);
                if (string.IsNullOrWhiteSpace(comparePath))
                    continue;

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
        if (tabs.Length == 0)
            return;

        var activeTab = tabs[0];
        for (var i = 0; i < tabs.Length; i++)
        {
            if (activeTab == tabHandle)
                break;

            SelectTabByIndex(windowHandle, i);
            activeTab = await Helper.DoUntilConditionAsync(
                () => WinApi.FindWindowEx(windowHandle, 0, "ShellTabWindowClass", null),
                handle => handle != activeTab);
        }
    }

    public void SelectTabByIndex(nint windowHandle, int index)
    {
        if (index < 0)
            return;

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
        if (tabHandle == 0)
            return;

        WinApi.PostMessage(tabHandle, WinApi.WM_COMMAND, 0xA21B, 0);

        if (bringToFront)
            WinApi.RestoreWindowToForeground(windowHandle);
    }

    private AutomationElement? GetExplorerTabItemAtPoint(Point screenPoint, nint explorerWindow)
    {
        try
        {
            var root = GetAutomationRoot(explorerWindow);
            if (root == null)
                return null;

            if (!IsPointInExplorerTabBand(screenPoint, explorerWindow))
                return null;

            var point = new System.Windows.Point(screenPoint.X, screenPoint.Y);
            var element = AutomationElement.FromPoint(point);
            var tabItem = GetExplorerTabTitleCandidate(element, explorerWindow, point);
            if (tabItem != null)
                return tabItem;

            var tabItems = root.FindAll(
                TreeScope.Descendants,
                new OrCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                    new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "tab item")));

            foreach (AutomationElement candidateTabItem in tabItems)
            {
                if (IsPointNearTabTitle(candidateTabItem, point) && IsUsableTabTitleBounds(candidateTabItem, explorerWindow) && !IsBlockedNonTabElement(candidateTabItem))
                    return candidateTabItem;
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private static bool IsPointInExplorerTabBand(Point screenPoint, nint explorerWindow)
    {
        if (!WinApi.GetWindowRect(explorerWindow, out var rect))
            return false;

        var topBandBottom = rect.Top + 120;
        return screenPoint.Y >= rect.Top && screenPoint.Y <= topBandBottom;
    }

    private static AutomationElement? GetExplorerTabTitleCandidate(AutomationElement? startingElement, nint explorerWindow, System.Windows.Point point)
    {
        if (startingElement == null)
            return null;

        var walker = TreeWalker.ControlViewWalker;
        var element = startingElement;

        while (element != null)
        {
            if (IsPointInsideElement(element, point) &&
                IsTabTitleLikeElement(element) &&
                IsUsableTabTitleBounds(element, explorerWindow))
                return element;

            if (IsBlockedNonTabElement(element))
                return null;

            var nativeHandle = element.Current.NativeWindowHandle;
            if (nativeHandle == explorerWindow)
                return null;

            element = walker.GetParent(element);
        }

        return null;
    }

    private static bool IsTabTitleLikeElement(AutomationElement element)
    {
        if (element.Current.ControlType == ControlType.TabItem)
            return true;

        var className = element.Current.ClassName ?? string.Empty;
        var automationId = element.Current.AutomationId ?? string.Empty;
        var localizedControlType = element.Current.LocalizedControlType ?? string.Empty;

        if (localizedControlType.Contains("tab item", StringComparison.OrdinalIgnoreCase))
            return true;

        if (automationId.Contains("Tab", StringComparison.OrdinalIgnoreCase) &&
            !automationId.Contains("TabBand", StringComparison.OrdinalIgnoreCase))
            return true;

        if (className.Contains("TabItem", StringComparison.OrdinalIgnoreCase))
            return true;

        return false;
    }

    private static bool IsBroadTabContainer(AutomationElement element, nint explorerWindow)
    {
        if (!WinApi.GetWindowRect(explorerWindow, out var windowRect))
            return true;

        var rect = element.Current.BoundingRectangle;
        if (rect.IsEmpty)
            return true;

        var windowWidth = Math.Max(1, windowRect.Right - windowRect.Left);
        return rect.Width > windowWidth * 0.45 || rect.Height > 64;
    }

    private static bool IsUsableTabTitleBounds(AutomationElement element, nint explorerWindow)
    {
        var rect = element.Current.BoundingRectangle;
        return !rect.IsEmpty && rect.Width >= 24 && rect.Height >= 12 && !IsBroadTabContainer(element, explorerWindow);
    }

    private static bool IsBlockedNonTabElement(AutomationElement element)
    {
        if (IsTabTitleLikeElement(element))
            return false;

        var className = element.Current.ClassName ?? string.Empty;
        var automationId = element.Current.AutomationId ?? string.Empty;
        var localizedControlType = element.Current.LocalizedControlType ?? string.Empty;
        var name = element.Current.Name ?? string.Empty;
        var controlType = element.Current.ControlType;

        if (controlType == ControlType.Edit ||
            controlType == ControlType.List ||
            controlType == ControlType.Tree ||
            controlType == ControlType.MenuBar ||
            controlType == ControlType.ToolBar ||
            controlType == ControlType.Button)
            return true;

        return className.Contains("Address", StringComparison.OrdinalIgnoreCase) ||
               className.Contains("Breadcrumb", StringComparison.OrdinalIgnoreCase) ||
               className.Contains("Search", StringComparison.OrdinalIgnoreCase) ||
               className.Contains("Toolbar", StringComparison.OrdinalIgnoreCase) ||
               automationId.Contains("Address", StringComparison.OrdinalIgnoreCase) ||
               automationId.Contains("Search", StringComparison.OrdinalIgnoreCase) ||
               automationId.Contains("Minimize", StringComparison.OrdinalIgnoreCase) ||
               automationId.Contains("Maximize", StringComparison.OrdinalIgnoreCase) ||
               automationId.Contains("Close", StringComparison.OrdinalIgnoreCase) ||
               localizedControlType.Contains("tool bar", StringComparison.OrdinalIgnoreCase) ||
               localizedControlType.Contains("edit", StringComparison.OrdinalIgnoreCase) ||
               localizedControlType.Contains("button", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("Address", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("Search", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("Minimize", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("Maximize", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("Close", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPointInsideElement(AutomationElement element, System.Windows.Point point)
    {
        var rect = element.Current.BoundingRectangle;
        return !rect.IsEmpty && rect.Contains(point);
    }

    private static bool IsPointNearTabTitle(AutomationElement element, System.Windows.Point point)
    {
        var rect = element.Current.BoundingRectangle;
        if (rect.IsEmpty)
            return false;

        return point.X >= rect.Left &&
               point.X <= rect.Right &&
               point.Y >= rect.Top &&
               point.Y <= rect.Bottom + 3;
    }

    private void StartExplorerProcessCheck() => _explorerCheckTimer = new Timer(CheckForMainExplorer, null, 0, 1000);

    private void CheckForMainExplorer(object? state)
    {
        var process = Helper.GetMainExplorerProcess();
        if (process == null)
            return;

        _explorerCheckTimer?.Dispose();
        _explorerCheckTimer = null;

        lock (_processLock)
        {
            if (_mainExplorerProcessId != 0)
                return;

            _mainExplorerProcessId = process.Id;
            InitializeShellObjects();
            StatusChanged?.Invoke("Explorer shell connected.");
            OnShellInitialized?.Invoke();
        }
    }

    private void OnExplorerProcessTerminated(object? sender, ProcessEventArgs e)
    {
        lock (_processLock)
        {
            if (e.ProcessId == _mainExplorerProcessId)
            {
                _mainExplorerProcessId = 0;
                DisposeShellObjects();
                StartExplorerProcessCheck();
                StatusChanged?.Invoke("Explorer restarted. Reconnecting hooks.");
                return;
            }
        }

        lock (_windowEntryDictLock)
        {
            for (var i = _windowEntryDict.Count - 1; i >= 0; i--)
            {
                var (window, info) = _windowEntryDict.ElementAt<WindowEntry>(i);
                try
                {
                    _ = window.HWND;
                }
                catch
                {
                    RemoveWindowAndUnhookEvents(window, info, useLock: false);
                }
            }
        }
    }

    private void InitializeShellObjects()
    {
        _shellPathComparer = new ShellPathComparer();
        _staTaskScheduler = new StaTaskScheduler();
        _shellWindows = new ShellWindows();
        _defaultLocation = Helper.GetDefaultExplorerLocation(_shellPathComparer);

        _windowRegisteredHandler = OnShellWindowRegistered;
        _shellWindows.WindowRegistered += _windowRegisteredHandler;

        foreach (var hWnd in Helper.GetAllExplorerWindows())
            _knownExplorerWindows.TryAdd(hWnd, 0);

        var count = _shellWindows.Count;
        for (var i = 0; i < count; i++)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;

            var windowInfo = new WindowInfo();
            lock (_windowEntryDictLock)
            {
                _windowEntryDict.Add(window, windowInfo);
            }

            window.PutProperty("seenBefore", true);
            _knownExplorerWindows.TryAdd(new IntPtr(window.HWND), 0);
            _ = GetTabHandle(window);
            HookWindowEvents(window, windowInfo);
        }

        _mainWindowHandle = GetMainWindowHWnd(0);
        _hiddenRecoveryTimer = new Timer(RecoverUnexpectedHiddenExplorerWindows, null, 1_000, 1_000);
    }

    private void DisposeShellObjects()
    {
        if (_shellWindows == null)
            return;

        if (_windowRegisteredHandler != null)
        {
            _shellWindows.WindowRegistered -= _windowRegisteredHandler;
            _windowRegisteredHandler = null;
        }

        _hiddenRecoveryTimer?.Dispose();
        _hiddenRecoveryTimer = null;

        foreach (var (window, info) in _windowEntryDict)
        {
            if (info.OnQuitHandler != null)
                window.OnQuit -= info.OnQuitHandler;

            if (info.OnNavigateHandler != null)
                window.NavigateComplete2 -= info.OnNavigateHandler;

            Marshal.ReleaseComObject(window);
        }

        _windowEntryDict.Clear();
        Marshal.ReleaseComObject(_shellWindows);
        _shellWindows = null;

        _shellPathComparer?.Dispose();
        _shellPathComparer = null;

        _staTaskScheduler?.Dispose();
        _staTaskScheduler = null;
        _mainWindowHandle = 0;
        _processedHWnds.Clear();
        _knownExplorerWindows.Clear();
        _mergeHiddenWindows.Clear();
        _stripBoundsCache.Clear();
        _stripBoundsRefreshInflight.Clear();
        _automationRootCache.Clear();
    }

    private void PreventWindowHiding(nint hWnd)
    {
        if (_processedHWnds.TryAdd(hWnd, 0))
            _ = Task.Delay(7_000).ContinueWith(_ => _processedHWnds.TryRemove(hWnd, out var removed), TaskScheduler.Default);
    }

    private bool IsPreHideSuppressed()
    {
        return Environment.TickCount64 < Interlocked.Read(ref _suppressPreHideUntil);
    }

    private InternetExplorer? GetRecentlyCreatedWindow(out WindowInfo? windowInfo)
    {
        if (_shellWindows == null)
        {
            windowInfo = null;
            return null;
        }

        var count = _shellWindows.Count;
        for (var i = count - 1; i >= 0; i--)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;

            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.Keys.Contains(window))
                    continue;

                try
                {
                    if (window.GetProperty("seenBefore") is not null)
                        continue;
                }
                catch
                {
                    continue;
                }

                try
                {
                    window.PutProperty("seenBefore", true);
                }
                catch
                {
                    continue;
                }

                windowInfo = new WindowInfo();
                if (!_windowEntryDict.TryAdd(window, windowInfo))
                    continue;

                try
                {
                    var hWnd = new IntPtr(window.HWND);
                    _knownExplorerWindows.TryAdd(hWnd, 0);
                    if (_mainWindowHandle == 0 && _windowEntryDict.Count == 1)
                        _mainWindowHandle = hWnd;
                }
                catch
                {
                    //
                }

                return window;
            }
        }

        windowInfo = null;
        return null;
    }

    private async void OnShellWindowRegistered(int unused)
    {
        var showAgain = true;
        nint hWnd = 0;

        try
        {
            var shouldOpenAsWindow = Helper.IsCtrlShiftDown();

            WindowInfo windowInfo = null!;
            var window = await Helper.DoUntilNotDefaultAsync(() => GetRecentlyCreatedWindow(out windowInfo!), 2_500, 70);
            if (window == null || windowInfo == null)
                return;

            _ = GetTabHandle(window);

            hWnd = new IntPtr(window.HWND);
            if (!Helper.IsFileExplorerWindow(hWnd))
            {
                HookWindowEvents(window, windowInfo);
                return;
            }

            if (shouldOpenAsWindow)
            {
                PreventWindowHiding(hWnd);
                HookWindowEvents(window, windowInfo);
                StatusChanged?.Invoke("Explorer window kept separate by Ctrl+Shift.");
                return;
            }

            var mainWindowHWnd = GetMainWindowHWnd(hWnd);
            var canReopenAsTabCandidate =
                (_isForcingTabs || _reuseTabs) &&
                IsMergeTargetWindow(mainWindowHWnd, hWnd) &&
                Helper.GetAllExplorerTabs(hWnd).Take(2).Count() == 1;

            if (canReopenAsTabCandidate)
                HideMergeCandidateWindow(hWnd);
            else
                PreventWindowHiding(hWnd);

            var location = await GetStableLocationAsync(window, windowInfo);
            var shouldReopenAsTab = canReopenAsTabCandidate && !IsProtectedLocation(location);

            if (shouldReopenAsTab)
            {
                showAgain = false;

                var merged = await OpenTabNavigateWithSelection(new WindowRecord(location, hWnd, GetSelectedItems(window), window.LocationName), mainWindowHWnd);
                if (merged)
                {
                    StatusChanged?.Invoke($"Merged Explorer window into a tab: {location}");
                    if (TryCloseMergedSourceWindow(window, hWnd, mainWindowHWnd))
                    {
                        _mergeHiddenWindows.TryRemove(hWnd, out _);
                        RemoveWindowAndUnhookEvents(window, windowInfo);
                        return;
                    }

                    showAgain = true;
                    StatusChanged?.Invoke("Merged tab created, but source Explorer window could not close. Restoring source window.");
                }

                showAgain = true;
                StatusChanged?.Invoke($"Merge validation failed. Left window open: {location}");
            }

            HookWindowEvents(window, windowInfo);
        }
        catch (Exception ex)
        {
            StatusChanged?.Invoke($"Explorer merge failed: {ex.Message}");
        }
        finally
        {
            if (showAgain && hWnd != 0)
            {
                ReleaseMergeHiddenWindow(hWnd, bringToFront: true);
            }
        }
    }

    private void HideMergeCandidateWindow(nint hWnd)
    {
        _mergeHiddenWindows[hWnd] = Environment.TickCount64;
        Helper.HideWindow(hWnd, keepTheme: true);
    }

    private void ReleaseMergeHiddenWindow(nint hWnd, bool bringToFront)
    {
        _mergeHiddenWindows.TryRemove(hWnd, out _);
        if (!Helper.IsFileExplorerWindow(hWnd))
        {
            Helper.HiddenWindows.TryRemove(hWnd, out _);
            return;
        }

        var wasHidden = Helper.ShowWindow(hWnd, removeCache: true);

        if (bringToFront)
            WinApi.RestoreWindowToForeground(hWnd);
        else if (wasHidden)
            WinApi.ShowWindow(hWnd, WinApi.SW_SHOWNOACTIVATE);
    }

    private void RecoverUnexpectedHiddenExplorerWindows(object? state)
    {
        if (Helper.HiddenWindows.IsEmpty)
            return;

        try
        {
            var now = Environment.TickCount64;
            var hasInteractiveExplorer = Helper.GetAllExplorerWindows().Any(IsExplorerWindowInteractive);
            foreach (var pair in Helper.HiddenWindows.ToArray())
            {
                var hWnd = pair.Key;
                if (!Helper.IsFileExplorerWindow(hWnd))
                {
                    Helper.HiddenWindows.TryRemove(hWnd, out _);
                    _mergeHiddenWindows.TryRemove(hWnd, out _);
                    continue;
                }

                if (_mergeHiddenWindows.TryGetValue(hWnd, out var hiddenAt))
                {
                    if (hasInteractiveExplorer && now - hiddenAt <= MergeHideGraceMs)
                        continue;

                    _mergeHiddenWindows.TryRemove(hWnd, out _);
                }

                ReleaseMergeHiddenWindow(hWnd, bringToFront: !hasInteractiveExplorer);
            }
        }
        catch
        {
            //
        }
    }

    private bool TryCloseMergedSourceWindow(InternetExplorer window, nint sourceHWnd, nint targetHWnd)
    {
        if (sourceHWnd == 0)
            return false;

        if (sourceHWnd == targetHWnd)
            return false;

        if (!Helper.IsFileExplorerWindow(sourceHWnd))
            return true;

        if (Helper.GetAllExplorerTabs(sourceHWnd).Take(2).Count() > 1)
            return false;

        try
        {
            window.Quit();
        }
        catch
        {
            return false;
        }

        var start = Stopwatch.GetTimestamp();
        while (!Helper.IsTimeUp(start, 1_500))
        {
            if (!Helper.IsFileExplorerWindow(sourceHWnd))
                return true;

            Thread.Sleep(60);
        }

        return !Helper.IsFileExplorerWindow(sourceHWnd);
    }

    private async Task<bool> OpenNewWindowWithSelection(WindowRecord windowToOpen, bool lockToOpenWindows = true)
    {
        if (lockToOpenWindows)
            await _toOpenWindowsLock.WaitAsync();

        try
        {
            nint[]? currentWindows = null;
            if (windowToOpen.SelectedItems?.Length > 0)
                currentWindows = Helper.GetAllExplorerWindows().ToArray();

            Helper.BypassWinForegroundRestrictions();

            var location = string.IsNullOrWhiteSpace(windowToOpen.Location) ? _defaultLocation : windowToOpen.Location;
            await RunInStaThread(() =>
            {
                Shell? shell = null;
                try
                {
                    shell = new Shell();
                    shell.ShellExecute(location, "", "", "opennewwindow");
                }
                finally
                {
                    if (shell != null)
                        Marshal.ReleaseComObject(shell);
                }
            });

            if (currentWindows == null)
                return true;

            var newWindowHandle = await Helper.ListenForNewExplorerWindowAsync(currentWindows);
            if (newWindowHandle == 0)
                return false;

            var window = _windowEntryDict.Keys.FirstOrDefault(w => w.HWND == newWindowHandle);
            if (window == null)
                return false;

            SelectItems(window, windowToOpen.SelectedItems);
            return true;
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
            if ((_reuseTabs || forceTabReuse) && !isDuplicate && _windowEntryDict.Count > 0)
            {
                var existingTab = SearchForTab(windowToOpen.Location);
                if (existingTab != 0)
                {
                    var targetWindow = WinApi.GetParent(existingTab);
                    await SelectTabByHandle(targetWindow, existingTab);
                    WinApi.RestoreWindowToForeground(targetWindow);
                    return true;
                }
            }

            var mainWindowHWnd = Helper.IsFileExplorerWindow(windowHandle)
                ? windowHandle
                : GetMainWindowHWnd(windowToOpen.Handle);

            if (mainWindowHWnd == 0)
                return await OpenNewWindowWithSelection(windowToOpen, lockToOpenWindows: false);

            var currentTabs = Helper.GetAllExplorerTabs(mainWindowHWnd).ToArray();
            await RequestToOpenNewTab(mainWindowHWnd, lockToOpenWindows: false);

            var newTabHandle = await Helper.ListenForNewExplorerTabAsync(mainWindowHWnd, currentTabs, 2_000);
            if (newTabHandle == 0)
                return false;

            var window = await Helper.DoUntilNotDefaultAsync(() => GetWindowByTabHandle(newTabHandle), 2_000, 50);
            if (window == null)
                return false;

            try
            {
                await Navigate(window, windowToOpen.Location);
                WinApi.RestoreWindowToForeground(mainWindowHWnd);

                if (!string.IsNullOrWhiteSpace(windowToOpen.Location))
                {
                    var finalLocation = await GetStableLocationAsync(window, _windowEntryDict[window].Value, 4_000);
                    if (!AreLocationsEquivalent(windowToOpen.Location, finalLocation))
                    {
                        WinApi.SendMessage(newTabHandle, WinApi.WM_COMMAND, 0xA021, 1);
                        return false;
                    }
                }

                SelectItems(window, windowToOpen.SelectedItems);
                return true;
            }
            catch
            {
                WinApi.SendMessage(newTabHandle, WinApi.WM_COMMAND, 0xA021, 1);
                return false;
            }
        }
        finally
        {
            _toOpenWindowsLock.Release();
        }
    }

    private async Task<string> GetStableLocationAsync(InternetExplorer window, WindowInfo windowInfo, int timeoutMs = 3_000)
    {
        string? latestLocation = null;
        var stableSamples = 0;
        var start = Stopwatch.GetTimestamp();

        while (!Helper.IsTimeUp(start, timeoutMs))
        {
            var current = TryGetLocation(window);
            if (string.IsNullOrWhiteSpace(current))
            {
                await Task.Delay(75);
                continue;
            }

            if (string.Equals(latestLocation, current, StringComparison.OrdinalIgnoreCase))
                stableSamples++;
            else
            {
                latestLocation = current;
                stableSamples = 0;
            }

            windowInfo.Location = current;
            windowInfo.Name = window.LocationName;

            if (!IsDefaultLocation(current) && stableSamples >= 1)
                return current;

            if (IsDefaultLocation(current) && stableSamples >= 6)
                return current;

            await Task.Delay(75);
        }

        return latestLocation ?? _defaultLocation;
    }

    private bool IsProtectedLocation(string location)
    {
        return location.StartsWith("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", StringComparison.OrdinalIgnoreCase);
    }

    private bool IsDefaultLocation(string location)
    {
        if (string.Equals(location, _defaultLocation, StringComparison.OrdinalIgnoreCase))
            return true;

        try
        {
            return _shellPathComparer != null && _shellPathComparer.IsEquivalent(location, _defaultLocation);
        }
        catch
        {
            return false;
        }
    }

    private bool AreLocationsEquivalent(string expected, string actual)
    {
        if (string.IsNullOrWhiteSpace(expected) || string.IsNullOrWhiteSpace(actual))
            return false;

        if (string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
            return true;

        try
        {
            return _shellPathComparer != null && _shellPathComparer.IsEquivalent(expected, actual);
        }
        catch
        {
            return false;
        }
    }

    private void HookWindowEvents(InternetExplorer window, WindowInfo windowInfo)
    {
        windowInfo.OnQuitHandler = () => RemoveWindowAndUnhookEvents(window, windowInfo);
        windowInfo.OnNavigateHandler = (object _, ref object _) =>
        {
            windowInfo.Location = TryGetLocation(window);
            windowInfo.Name = window.LocationName;
        };

        try
        {
            window.OnQuit += windowInfo.OnQuitHandler;
            window.NavigateComplete2 += windowInfo.OnNavigateHandler;

            windowInfo.Location = TryGetLocation(window);
            windowInfo.Name = window.LocationName;

            // Eagerly compute the tab strip bounding rect so the very first double-click on a tab in this
            // window hits the cache (no UIA on the mouse hook thread, no "click many times to make it work").
            try
            {
                var hWnd = new IntPtr(window.HWND);
                if (Helper.IsFileExplorerWindow(hWnd))
                    ScheduleStripBoundsRefresh(hWnd);
            }
            catch
            {
                //
            }
        }
        catch
        {
            lock (_windowEntryDictLock)
            {
                _windowEntryDict.Remove(window);
            }
        }
    }

    private void RemoveWindowAndUnhookEvents(InternetExplorer window, WindowInfo windowInfo, bool useLock = true)
    {
        if (windowInfo.OnQuitHandler != null)
            window.OnQuit -= windowInfo.OnQuitHandler;

        if (windowInfo.OnNavigateHandler != null)
            window.NavigateComplete2 -= windowInfo.OnNavigateHandler;

        if (useLock)
        {
            lock (_windowEntryDictLock)
            {
                _windowEntryDict.Remove(window);
            }
        }
        else
        {
            _windowEntryDict.Remove(window);
        }

        try
        {
            var windowHandle = new IntPtr(window.HWND);
            bool stillTrackedByHandle;
            lock (_windowEntryDictLock)
            {
                stillTrackedByHandle = IsWindowTrackedByHandle(windowHandle);
            }

            if (!stillTrackedByHandle)
            {
                var windowStillExists = Helper.IsFileExplorerWindow(windowHandle);
                _processedHWnds.TryRemove(windowHandle, out _);

                if (!windowStillExists)
                {
                    Helper.HiddenWindows.TryRemove(windowHandle, out _);
                    _mergeHiddenWindows.TryRemove(windowHandle, out _);
                    _knownExplorerWindows.TryRemove(windowHandle, out _);
                    InvalidateAutomationRoot(windowHandle);
                    InvalidateStripBounds(windowHandle);
                    if (_mainWindowHandle == windowHandle)
                        _mainWindowHandle = 0;
                }
            }
        }
        catch
        {
            //
        }

        Marshal.ReleaseComObject(window);
    }

    private nint GetMainWindowHWnd(nint otherThan)
    {
        if (IsMergeTargetWindow(_mainWindowHandle, otherThan))
            return _mainWindowHandle;

        var foreground = WinApi.GetForegroundWindow();
        var allWindows = WinApi.FindAllWindowsEx("CabinetWClass");
        _mainWindowHandle = allWindows
            .Where(handle => IsMergeTargetWindow(handle, otherThan))
            .OrderByDescending(handle => handle == foreground)
            .ThenByDescending(WinApi.IsWindowVisible)
            .ThenByDescending(handle => WinApi.FindAllWindowsEx("ShellTabWindowClass", handle).Count())
            .FirstOrDefault();

        return _mainWindowHandle;
    }

    private bool HasLiveMergeTarget(nint otherThan)
    {
        var mainWindow = GetMainWindowHWnd(otherThan);
        return mainWindow != 0 &&
               mainWindow != otherThan &&
               Helper.IsFileExplorerWindow(mainWindow) &&
               Helper.GetAllExplorerTabs(mainWindow).Any();
    }

    private bool IsWindowTrackedByHandle(nint hWnd)
    {
        foreach (var (window, _) in _windowEntryDict)
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

        return false;
    }

    private bool IsMergeTargetWindow(nint hWnd, nint otherThan = 0)
    {
        if (hWnd == 0 || hWnd == otherThan || !Helper.IsFileExplorerWindow(hWnd))
            return false;

        if (_mergeHiddenWindows.ContainsKey(hWnd) || Helper.HiddenWindows.ContainsKey(hWnd))
            return false;

        return Helper.GetAllExplorerTabs(hWnd).Any();
    }

    private bool IsExplorerWindowInteractive(nint hWnd)
    {
        if (!Helper.IsFileExplorerWindow(hWnd))
            return false;

        if (_mergeHiddenWindows.ContainsKey(hWnd) || Helper.HiddenWindows.ContainsKey(hWnd))
            return false;

        return WinApi.IsWindowVisible(hWnd) && !WinApi.IsIconic(hWnd);
    }

    private Task<nint> GetTabHandle(InternetExplorer window)
    {
        if (_windowEntryDict.TryGetValue(window, out WindowEntry entry) && entry.OptionalKey is { } handle && handle > 0)
            return Task.FromResult(handle);

        return RunInStaThread(() =>
        {
            if (window is not WinTab.Interop.IServiceProvider serviceProvider)
                return 0;

            serviceProvider.QueryService(ref ShellBrowserGuid, ref ShellBrowserGuid, out var shellBrowser);
            if (shellBrowser == null)
                return 0;

            try
            {
                shellBrowser.GetWindow(out var handle);
                if (handle != 0)
                    _windowEntryDict.UpdateOptionalKey(window, handle);

                return handle;
            }
            finally
            {
                Marshal.ReleaseComObject(shellBrowser);
            }
        });
    }

    private static nint GetActiveTabHandle(nint windowHandle)
    {
        return WinApi.FindWindowEx(windowHandle, 0, "ShellTabWindowClass", null);
    }

    private InternetExplorer? GetWindowByTabHandle(nint tabHandle)
    {
        if (tabHandle == 0)
            return null;

        return _windowEntryDict.TryGetValue(tabHandle, out InternetExplorer? foundWindow) ? foundWindow : null;
    }

    private static string[]? GetSelectedItems(InternetExplorer window)
    {
        if (window.Document is not ShellFolderView document)
            return null;

        var selectedItems = document.SelectedItems();
        var count = selectedItems.Count;
        if (count == 0)
            return null;

        var result = new string[count];
        for (var i = 0; i < count; i++)
            result[i] = selectedItems.Item(i).Name;

        return result;
    }

    private static void SelectItems(InternetExplorer window, string[]? names)
    {
        if (names == null || names.Length == 0 || window.Document is not ShellFolderView document)
            return;

        for (var i = 0; i < names.Length; i++)
        {
            object item = document.Folder.ParseName(names[i]);
            if (item == null)
                continue;

            document.SelectItem(ref item, 1);
        }
    }

    private string TryGetLocation(InternetExplorer window)
    {
        try
        {
            var path = window.LocationURL;
            if (!string.IsNullOrWhiteSpace(path))
                return Helper.NormalizeLocation(path);

            path = ((window.Document as ShellFolderView)?.Folder as Folder2)?.Self.Path;
            return string.IsNullOrWhiteSpace(path) ? string.Empty : Helper.NormalizeLocation(path);
        }
        catch
        {
            return string.Empty;
        }
    }

    private async Task Navigate(InternetExplorer window, string path)
    {
        if (!path.Contains("#", StringComparison.Ordinal) && !path.Contains("%23", StringComparison.OrdinalIgnoreCase))
        {
            window.Navigate2(path);
            return;
        }

        var folder = await RunInStaThread(() =>
        {
            Shell? shell = null;
            try
            {
                shell = new Shell();
                return shell.NameSpace(path);
            }
            finally
            {
                if (shell != null)
                    Marshal.ReleaseComObject(shell);
            }
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

    private Task RunInStaThread(Action action, CancellationToken cancellationToken = default)
    {
        return Task.Factory.StartNew(action, cancellationToken, TaskCreationOptions.None, _staTaskScheduler!);
    }

    private Task<T> RunInStaThread<T>(Func<T> action, CancellationToken cancellationToken = default)
    {
        return Task.Factory.StartNew(action, cancellationToken, TaskCreationOptions.None, _staTaskScheduler!);
    }

    public void Dispose()
    {
        DisposeShellObjects();
        _processWatcher.Dispose();
        _toOpenWindowsLock.Dispose();
        _instanceRunning = false;
        GC.SuppressFinalize(this);
    }
}
