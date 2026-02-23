// Adapted from ExplorerTabUtility (MIT License).
// Source: E:\_BIGFA Free\_code\ExplorerTabUtility

using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Threading;
using WinTab.App.ExplorerTabUtilityPort.Interop;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;
using SendKeys = System.Windows.Forms.SendKeys;

namespace WinTab.App.ExplorerTabUtilityPort;

/// <summary>
/// Event-driven Explorer window-to-tab conversion.
/// Uses WinEvent hooks for new windows and COM (Shell.Application) to
/// navigate the correct tab object (avoids leaving a "This PC" tab behind).
/// </summary>
public sealed class ExplorerTabHookService : IDisposable
{
    public async Task<bool> OpenLocationAsTabAsync(string location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return false;

        location = location.Trim();

        if (await TryNavigateCurrentActiveTabLikeExplorer(location))
            return true;

        if (await TryActivateExistingTabByLocation(location))
            return true;

        // If no Explorer window is available, fall back to normal explorer.exe open.
        IntPtr targetTopLevel = PickTargetExplorerWindow(exclude: IntPtr.Zero);
        if (targetTopLevel == IntPtr.Zero)
        {
            return TryLaunchExplorerWindow(location);
        }

        bool opened = await OpenLocationInNewTab(targetTopLevel, location);
        if (!opened)
        {
            _logger.Warn($"OpenLocationAsTabAsync: tab-open failed, fallback to normal explorer open for '{location}'.");
            return TryLaunchExplorerWindow(location);
        }

        TryBringToForeground(targetTopLevel);
        return true;
    }

    private bool TryLaunchExplorerWindow(string location)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{location}\"",
                UseShellExecute = true
            });
            return true;
        }
        catch (Exception ex)
        {
            _logger.Warn($"Failed to launch explorer fallback for '{location}': {ex.Message}");
            return false;
        }
    }

    private async Task<bool> TryNavigateCurrentActiveTabLikeExplorer(string location)
    {
        if (_settings.OpenChildFolderInNewTabFromActiveTab)
            return false;

        if (!IsRealFileSystemLocation(location))
            return false;

        IntPtr foreground = NativeMethods.GetForegroundWindow();
        if (foreground == IntPtr.Zero || !IsExplorerTopLevelWindow(foreground))
            return false;

        IntPtr activeTab = GetActiveTabHandle(foreground);
        if (activeTab == IntPtr.Zero || !NativeMethods.IsWindow(activeTab))
            return false;

        string? currentLocation = await UiAsync(() => TryGetLocationByTabHandleUi(activeTab));
        if (!IsRealFileSystemLocation(currentLocation))
            return false;

        if (!IsChildPathOf(currentLocation!, location))
            return false;

        bool navigated = await NavigateTabByHandleWithRetry(activeTab, location, timeoutMs: 420);
        if (!navigated)
            return false;

        bool confirmed = await WaitUntilTabLocationMatches(activeTab, location, timeoutMs: 420, pollMs: 40);
        if (!confirmed)
            return false;

        TryBringToForeground(foreground);
        _logger.Info($"Navigated active Explorer tab directly: {location}");
        return true;
    }

    private static readonly Guid ShellBrowserGuid = typeof(IShellBrowser).GUID;
    private static readonly Guid ShellWindowsClsid = new("9BA05972-F6A8-11CF-A442-00A0C90A8F39");
    private static readonly Guid ShellWindowsEventsGuid = new("FE4106E0-399A-11D0-A48C-00A0C90A8F39");
    private const string ExplorerExe = "explorer.exe";
    private const string ExplorerTabClass = "ShellTabWindowClass";
    private const string ExplorerWindowClass = "CabinetWClass";

    private static readonly object ShellWindowsInitLock = new();
    private static object? _shellWindows;
    private static int _shellWindowsThreadId;

    private readonly IWindowEventSource _windowEvents;
    private readonly IWindowManager _windowManager;
    private readonly AppSettings _settings;
    private readonly Logger _logger;

    private readonly Action<int> _windowRegisteredHandler;
    private readonly NativeMethods.WinEventDelegate _createEventCallback;
    private IConnectionPointNative? _shellWindowsConnectionPoint;
    private object? _shellWindowsEventsSink;
    private int _shellWindowsConnectionCookie;
    private bool _shellWindowRegisteredHooked;
    private readonly IntPtr _createEventHook;

    private readonly ConcurrentDictionary<IntPtr, DateTimeOffset> _pending = new();
    private readonly ConcurrentDictionary<IntPtr, byte> _knownExplorerTopLevels = new();
    private readonly ConcurrentDictionary<IntPtr, byte> _earlyHiddenExplorer = new();
    private readonly ConcurrentDictionary<IntPtr, ConcurrentDictionary<IntPtr, int>> _tabIndexCache = new();
    private readonly object _sendLock = new();

    private IntPtr _lastExplorerForeground;

    private bool _disposed;

    public ExplorerTabHookService(IWindowEventSource windowEvents, IWindowManager windowManager, AppSettings settings, Logger logger)
    {
        _windowEvents = windowEvents;
        _windowManager = windowManager;
        _settings = settings;
        _logger = logger;
        _windowRegisteredHandler = OnShellWindowRegistered;
        _createEventCallback = OnExplorerObjectCreate;

        bool enableAutoConvert = string.Equals(
            Environment.GetEnvironmentVariable("WINTAB_AUTO_CONVERT_EXPLORER"),
            "1",
            StringComparison.Ordinal);

        if (!enableAutoConvert)
        {
            _createEventHook = IntPtr.Zero;
            _logger.Info("ExplorerTabHookService started in on-demand mode (auto-convert disabled).");
            return;
        }

        _windowEvents.WindowShown += OnWindowShown;
        _windowEvents.WindowForegroundChanged += OnWindowForegroundChanged;
        _windowEvents.WindowDestroyed += OnWindowDestroyed;

        IntPtr currentForeground = NativeMethods.GetForegroundWindow();
        if (currentForeground != IntPtr.Zero && IsExplorerTopLevelWindow(currentForeground))
            _lastExplorerForeground = currentForeground;

        // Seed cache with existing Explorer windows so we don't treat normal
        // "show" events (navigation, layout changes) as new windows.
        foreach (IntPtr h in EnumerateExplorerTopLevelWindows(includeInvisible: true))
            _knownExplorerTopLevels.TryAdd(h, 0);

        _shellWindowRegisteredHooked = TryHookShellWindowRegistered();
        if (_shellWindowRegisteredHooked)
        {
            _logger.Info("ExplorerTabHookService: ShellWindows.WindowRegistered hooked.");
        }
        else
        {
            _logger.Warn("ExplorerTabHookService: ShellWindows.WindowRegistered hook unavailable, using WindowShown fallback.");
        }

        // Always install EVENT_OBJECT_CREATE to reduce flash when possible.
        const uint flags = NativeConstants.WINEVENT_OUTOFCONTEXT | NativeConstants.WINEVENT_SKIPOWNPROCESS;
        _createEventHook = NativeMethods.SetWinEventHook(
            NativeConstants.EVENT_OBJECT_CREATE,
            NativeConstants.EVENT_OBJECT_CREATE,
            IntPtr.Zero,
            _createEventCallback,
            0,
            0,
            flags);

        if (_createEventHook == IntPtr.Zero)
            _logger.Warn("ExplorerTabHookService: EVENT_OBJECT_CREATE hook unavailable (flash may occur).");

        _logger.Info("ExplorerTabHookService started (ETU-style, no COMReference).");
    }

    private readonly record struct RegisteredCandidate(IntPtr TopLevelHwnd, IntPtr TabHandle, string? Location);
    private readonly record struct ExistingTabCandidate(IntPtr TopLevelHwnd, IntPtr TabHandle, string Location);

    private async Task<bool> TryActivateExistingTabByLocation(string location, IntPtr excludeTopLevel = default)
    {
        ExistingTabCandidate? candidate = await UiAsync(() => TryFindExistingTabByLocationUi(location, excludeTopLevel));
        if (candidate is null)
            return false;

        bool activated = await TryActivateTabHandle(candidate.Value.TopLevelHwnd, candidate.Value.TabHandle);
        if (!activated)
        {
            // Do not create duplicate tabs when an existing tab has already been found.
            TryBringToForeground(candidate.Value.TopLevelHwnd);
            _logger.Warn($"Found existing tab for '{location}' but failed to activate it precisely; skipped duplicate creation.");
            return true;
        }

        _logger.Info($"Activated existing Explorer tab: {candidate.Value.Location}");
        return true;
    }

    private ExistingTabCandidate? TryFindExistingTabByLocationUi(string location, IntPtr excludeTopLevel)
    {
        if (!IsRealFileSystemLocation(location))
            return null;

        IntPtr foreground = NativeMethods.GetForegroundWindow();
        IntPtr lastForeground = _lastExplorerForeground;

        ExistingTabCandidate? firstMatch = null;
        ExistingTabCandidate? preferredLastForegroundMatch = null;

        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                string? tabLocation = TryGetComLocation(tab);
                if (tabLocation is null ||
                    !IsRealFileSystemLocation(tabLocation) ||
                    !PathsEquivalent(tabLocation, location))
                    continue;

                IntPtr tabHandle = GetTabHandle(tab);
                if (tabHandle == IntPtr.Zero || !NativeMethods.IsWindow(tabHandle))
                    continue;

                IntPtr topLevel = NativeMethods.GetAncestor(tabHandle, NativeConstants.GA_ROOT);
                if (topLevel == IntPtr.Zero || !NativeMethods.IsWindow(topLevel))
                    continue;

                if (excludeTopLevel != IntPtr.Zero && topLevel == excludeTopLevel)
                    continue;

                if (!IsExplorerTopLevelWindow(topLevel))
                    continue;

                var candidate = new ExistingTabCandidate(topLevel, tabHandle, tabLocation!);
                if (topLevel == foreground)
                    return candidate;

                if (lastForeground != IntPtr.Zero && topLevel == lastForeground)
                    preferredLastForegroundMatch = candidate;

                firstMatch ??= candidate;
            }
            catch
            {
                // ignore
            }
            finally
            {
                Marshal.FinalReleaseComObject(tab);
            }
        }

        return preferredLastForegroundMatch ?? firstMatch;
    }

    private async Task<bool> TryActivateTabHandle(IntPtr explorerTopLevel, IntPtr tabHandle)
    {
        if (explorerTopLevel == IntPtr.Zero || tabHandle == IntPtr.Zero)
            return false;

        if (!NativeMethods.IsWindow(explorerTopLevel) || !NativeMethods.IsWindow(tabHandle))
            return false;

        bool targetIsForeground = NativeMethods.GetForegroundWindow() == explorerTopLevel;

        if (!targetIsForeground)
            NativeMethods.ShowWindow(explorerTopLevel, NativeConstants.SW_SHOWNOACTIVATE);

        if (GetActiveTabHandle(explorerTopLevel) == tabHandle)
        {
            TryBringToForeground(explorerTopLevel);
            return true;
        }

        List<IntPtr> tabs = GetAllTabHandles(explorerTopLevel);
        if (tabs.Count == 0)
            return false;

        int preferredIndex = TryGetCachedTabIndex(explorerTopLevel, tabHandle, tabs);
        if (preferredIndex >= 0 && await TryActivateTabByIndex(explorerTopLevel, tabHandle, preferredIndex, timeoutMs: 220))
        {
            RememberTabIndex(explorerTopLevel, tabHandle, preferredIndex);
            TryBringToForeground(explorerTopLevel);
            return true;
        }

        int comIndex = await UiAsync(() => TryGetComTabIndexUi(explorerTopLevel, tabHandle));

        if (comIndex >= 0 &&
            comIndex != preferredIndex &&
            await TryActivateTabByIndex(explorerTopLevel, tabHandle, comIndex, timeoutMs: 220))
        {
            RememberTabIndex(explorerTopLevel, tabHandle, comIndex);
            TryBringToForeground(explorerTopLevel);
            return true;
        }

        int matchedIndex;
        if (targetIsForeground)
        {
            bool lockReady = TryEnterFrontProbeRedrawLock(explorerTopLevel, out bool redrawLocked, out bool windowLocked);
            if (!lockReady)
            {
                _logger.Warn($"Activation probe skipped (no redraw lock) for Explorer 0x{explorerTopLevel.ToInt64():X}.");
                return false;
            }

            try
            {
                matchedIndex = await ProbeTabIndexForHandle(explorerTopLevel, tabHandle, tabs, skipPrimary: preferredIndex, skipSecondary: comIndex);
            }
            finally
            {
                ExitFrontProbeRedrawLock(explorerTopLevel, redrawLocked, windowLocked);
            }
        }
        else
        {
            matchedIndex = await ProbeTabIndexForHandle(explorerTopLevel, tabHandle, tabs, skipPrimary: preferredIndex, skipSecondary: comIndex);
        }

        if (matchedIndex >= 0)
        {
            RememberTabIndex(explorerTopLevel, tabHandle, matchedIndex);
            TryBringToForeground(explorerTopLevel);
            return true;
        }

        bool activeNow = GetActiveTabHandle(explorerTopLevel) == tabHandle;
        if (activeNow)
            TryBringToForeground(explorerTopLevel);

        return activeNow;
    }

    private int TryGetCachedTabIndex(IntPtr explorerTopLevel, IntPtr tabHandle, List<IntPtr> tabs)
    {
        if (_tabIndexCache.TryGetValue(explorerTopLevel, out ConcurrentDictionary<IntPtr, int>? perWindow) &&
            perWindow.TryGetValue(tabHandle, out int cachedIndex) &&
            cachedIndex >= 0 &&
            cachedIndex < tabs.Count)
        {
            return cachedIndex;
        }

        return -1;
    }

    private void RememberTabIndex(IntPtr explorerTopLevel, IntPtr tabHandle, int index)
    {
        if (index < 0)
            return;

        ConcurrentDictionary<IntPtr, int> perWindow = _tabIndexCache.GetOrAdd(
            explorerTopLevel,
            static _ => new ConcurrentDictionary<IntPtr, int>());

        perWindow[tabHandle] = index;
    }

    private async Task<bool> TryActivateTabByIndex(IntPtr explorerTopLevel, IntPtr tabHandle, int index, int timeoutMs)
    {
        if (!TrySelectTabByIndex(explorerTopLevel, index))
            return false;

        return await WaitForActiveTabHandle(explorerTopLevel, tabHandle, timeoutMs: timeoutMs);
    }

    private async Task<int> ProbeTabIndexForHandle(IntPtr explorerTopLevel, IntPtr tabHandle, List<IntPtr> tabs, int skipPrimary, int skipSecondary)
    {
        for (int i = 0; i < tabs.Count; i++)
        {
            if (i == skipPrimary || i == skipSecondary)
                continue;

            if (!TrySelectTabByIndex(explorerTopLevel, i))
                continue;

            if (await WaitForActiveTabHandle(explorerTopLevel, tabHandle, timeoutMs: 120))
                return i;
        }

        return -1;
    }

    private static bool TryEnterFrontProbeRedrawLock(IntPtr explorerTopLevel, out bool redrawLocked, out bool windowLocked)
    {
        redrawLocked = false;
        windowLocked = false;

        if (explorerTopLevel == IntPtr.Zero || !NativeMethods.IsWindow(explorerTopLevel))
            return false;

        try
        {
            windowLocked = NativeMethods.LockWindowUpdate(explorerTopLevel);
        }
        catch
        {
            // ignore
        }

        if (!windowLocked)
            return false;

        try
        {
            NativeMethods.SendMessage(
                explorerTopLevel,
                (uint)NativeConstants.WM_SETREDRAW,
                IntPtr.Zero,
                IntPtr.Zero);

            redrawLocked = true;
        }
        catch
        {
            // ignore
        }

        return true;
    }

    private static void ExitFrontProbeRedrawLock(IntPtr explorerTopLevel, bool redrawLocked, bool windowLocked)
    {
        if (windowLocked)
        {
            try
            {
                NativeMethods.LockWindowUpdate(IntPtr.Zero);
            }
            catch
            {
                // ignore
            }
        }

        if (!redrawLocked || explorerTopLevel == IntPtr.Zero || !NativeMethods.IsWindow(explorerTopLevel))
            return;

        try
        {
            NativeMethods.SendMessage(
                explorerTopLevel,
                (uint)NativeConstants.WM_SETREDRAW,
                new IntPtr(1),
                IntPtr.Zero);

            uint flags = NativeConstants.RDW_INVALIDATE |
                         NativeConstants.RDW_ERASE |
                         NativeConstants.RDW_FRAME |
                         NativeConstants.RDW_ALLCHILDREN |
                         NativeConstants.RDW_UPDATENOW;

            NativeMethods.RedrawWindow(explorerTopLevel, IntPtr.Zero, IntPtr.Zero, flags);
        }
        catch
        {
            // ignore
        }
    }

    private static int TryGetComTabIndexUi(IntPtr explorerTopLevel, IntPtr tabHandle)
    {
        if (explorerTopLevel == IntPtr.Zero || tabHandle == IntPtr.Zero)
            return -1;

        int index = 0;
        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                IntPtr handle = GetTabHandle(tab);
                if (handle == IntPtr.Zero)
                    continue;

                IntPtr topLevel = NativeMethods.GetAncestor(handle, NativeConstants.GA_ROOT);
                if (topLevel != explorerTopLevel)
                    continue;

                if (handle == tabHandle)
                    return index;

                index++;
            }
            catch
            {
                // ignore
            }
            finally
            {
                Marshal.FinalReleaseComObject(tab);
            }
        }

        return -1;
    }

    private static bool TrySelectTabByIndex(IntPtr explorerTopLevel, int index)
    {
        if (index < 0 || explorerTopLevel == IntPtr.Zero || !NativeMethods.IsWindow(explorerTopLevel))
            return false;

        NativeMethods.SendMessage(
            explorerTopLevel,
            (uint)NativeConstants.WM_COMMAND,
            new IntPtr(NativeConstants.EXPLORER_CMD_SELECT_TAB_BY_INDEX),
            new IntPtr(index + 1));

        return true;
    }

    private static async Task<bool> WaitForActiveTabHandle(IntPtr explorerTopLevel, IntPtr expectedTabHandle, int timeoutMs)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            if (GetActiveTabHandle(explorerTopLevel) == expectedTabHandle)
                return true;

            await Task.Delay(25);
        }

        return GetActiveTabHandle(explorerTopLevel) == expectedTabHandle;
    }

    private bool TryHookShellWindowRegistered()
    {
        try
        {
            object? windows = UiAsync(GetShellWindowsCollectionUi).GetAwaiter().GetResult();
            if (windows is null)
                return false;

            if (windows is not IConnectionPointContainerNative cpc)
                return false;

            Guid iid = ShellWindowsEventsGuid;
            cpc.FindConnectionPoint(ref iid, out IConnectionPointNative cp);
            if (cp is null)
                return false;

            var sink = new ShellWindowsEventsSink(_windowRegisteredHandler);
            cp.Advise(sink, out int cookie);
            if (cookie == 0)
            {
                Marshal.FinalReleaseComObject(cp);
                return false;
            }

            _shellWindowsConnectionPoint = cp;
            _shellWindowsEventsSink = sink;
            _shellWindowsConnectionCookie = cookie;

            return true;
        }
        catch (Exception ex)
        {
            _logger.Warn($"ExplorerTabHookService: WindowRegistered hook error: {ex.GetType().Name}: {ex.Message}");
            return false;
        }
    }

    private async void OnShellWindowRegistered(int cookie)
    {
        if (_disposed)
            return;

        try
        {
            RegisteredCandidate? candidate = await WaitForRegisteredCandidate(cookie, timeoutMs: 1200);
            if (candidate is null)
                return;

            IntPtr hwnd = candidate.Value.TopLevelHwnd;
            if (!_pending.TryAdd(hwnd, DateTimeOffset.UtcNow))
                return;

            _logger.Info($"Explorer candidate window registered: 0x{hwnd.ToInt64():X}");

            bool hiddenImmediately = _earlyHiddenExplorer.TryRemove(hwnd, out _);
            if (!hiddenImmediately)
                hiddenImmediately = _windowManager.Hide(hwnd);

            _ = Task.Run(async () =>
            {
                try
                {
                    await ConvertWindowToTab(
                        hwnd,
                        candidate.Value.TabHandle,
                        candidate.Value.Location,
                        sourceHiddenAlready: hiddenImmediately);
                }
                catch (Exception ex)
                {
                    _logger.Error($"Explorer tab hook failed for {hwnd}.", ex);
                }
                finally
                {
                    _pending.TryRemove(hwnd, out _);
                }
            });
        }
        catch
        {
            // ignore and keep fallback path alive
        }
    }

    private void OnExplorerObjectCreate(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr hwnd,
        int idObject,
        int idChild,
        uint dwEventThread,
        uint dwmsEventTime)
    {
        if (_disposed || hwnd == IntPtr.Zero)
            return;

        if (idObject != NativeConstants.OBJID_WINDOW || idChild != 0)
            return;

        if (_pending.ContainsKey(hwnd) || _knownExplorerTopLevels.ContainsKey(hwnd))
            return;

        var classBuilder = new StringBuilder(32);
        NativeMethods.GetClassName(hwnd, classBuilder, classBuilder.Capacity);
        if (!string.Equals(classBuilder.ToString(), ExplorerWindowClass, StringComparison.OrdinalIgnoreCase))
            return;

        // Hide first (fast) to reduce flash; we'll validate later.
        // Only do this when the top-level window class matches.
        if (_windowManager.Hide(hwnd))
            _earlyHiddenExplorer.TryAdd(hwnd, 0);
    }

    private async Task<RegisteredCandidate?> WaitForRegisteredCandidate(int cookie, int timeoutMs)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            RegisteredCandidate? candidate = await UiAsync(() =>
                TryTakeRegisteredCandidateByCookieUi(cookie) ?? TryTakeRegisteredCandidateUi());

            if (candidate is not null)
                return candidate;

            await Task.Delay(25);
        }

        return null;
    }

    private RegisteredCandidate? TryTakeRegisteredCandidateByCookieUi(int cookie)
    {
        object? windows = GetShellWindowsCollectionUi();
        if (windows is null)
            return null;

        object? tab = null;
        try
        {
            dynamic dynWindows = windows;
            tab = dynWindows.Item(cookie);
            if (tab is null)
                return null;

            dynamic win = tab;
            IntPtr topLevel = new IntPtr((int)win.HWND);
            if (topLevel == IntPtr.Zero)
                return null;

            if (_knownExplorerTopLevels.ContainsKey(topLevel))
                return null;

            if (!IsExplorerTopLevelWindow(topLevel))
                return null;

            _knownExplorerTopLevels.TryAdd(topLevel, 0);

            IntPtr tabHandle = GetTabHandle(tab);
            string? location = TryGetComLocation(tab);

            return new RegisteredCandidate(topLevel, tabHandle, location);
        }
        catch
        {
            return null;
        }
        finally
        {
            if (tab is not null)
                Marshal.FinalReleaseComObject(tab);
        }
    }

    private RegisteredCandidate? TryTakeRegisteredCandidateUi()
    {
        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                dynamic win = tab;
                IntPtr topLevel = new IntPtr((int)win.HWND);
                if (topLevel == IntPtr.Zero)
                    continue;

                if (_knownExplorerTopLevels.ContainsKey(topLevel))
                    continue;

                if (!IsExplorerTopLevelWindow(topLevel))
                    continue;

                _knownExplorerTopLevels.TryAdd(topLevel, 0);

                IntPtr tabHandle = GetTabHandle(tab);
                string? location = TryGetComLocation(tab);

                return new RegisteredCandidate(topLevel, tabHandle, location);
            }
            catch
            {
                // ignore
            }
            finally
            {
                Marshal.FinalReleaseComObject(tab);
            }
        }

        return null;
    }

    private void OnWindowForegroundChanged(object? sender, IntPtr hwnd)
    {
        if (_disposed || hwnd == IntPtr.Zero)
            return;

        if (IsExplorerTopLevelWindow(hwnd))
            _lastExplorerForeground = hwnd;
    }

    private static Dispatcher? TryGetUiDispatcher()
    {
        try
        {
            return System.Windows.Application.Current?.Dispatcher;
        }
        catch
        {
            return null;
        }
    }

    private static Task<T> UiAsync<T>(Func<T> func)
    {
        Dispatcher? dispatcher = TryGetUiDispatcher();
        if (dispatcher is null)
            return Task.FromResult(func());

        if (dispatcher.CheckAccess())
            return Task.FromResult(func());

        return dispatcher.InvokeAsync(func).Task;
    }

    private void OnWindowDestroyed(object? sender, IntPtr hwnd)
    {
        if (_disposed || hwnd == IntPtr.Zero)
            return;

        _knownExplorerTopLevels.TryRemove(hwnd, out _);
        _earlyHiddenExplorer.TryRemove(hwnd, out _);
        _pending.TryRemove(hwnd, out _);
        _tabIndexCache.TryRemove(hwnd, out _);
    }

    private void OnWindowShown(object? sender, IntPtr hwnd)
    {
        if (_disposed || hwnd == IntPtr.Zero)
            return;

        // Even when WindowRegistered is hooked, SHOW can still arrive earlier or be used for re-hide.
        TryQueueShownCandidate(hwnd);
    }

    private void TryQueueShownCandidate(IntPtr hwnd)
    {
        if (_pending.ContainsKey(hwnd))
            return;

        // Fast filter: only handle visible Explorer windows.
        if (!NativeMethods.IsWindowVisible(hwnd))
            return;

        if (!IsExplorerTopLevelWindow(hwnd))
            return;

        // Ignore repeated show events for already-known windows.
        if (!_knownExplorerTopLevels.TryAdd(hwnd, 0))
            return;

        _logger.Info($"Explorer candidate window shown: 0x{hwnd.ToInt64():X}");

        // Prevent re-entrancy for the same hwnd.
        if (!_pending.TryAdd(hwnd, DateTimeOffset.UtcNow))
            return;

        _ = Task.Run(async () =>
        {
            try
            {
                await ConvertWindowToTab(hwnd);
            }
            catch (Exception ex)
            {
                _logger.Error($"Explorer tab hook failed for {hwnd}.", ex);
            }
            finally
            {
                _pending.TryRemove(hwnd, out _);
            }
        });
    }

    private bool IsExplorerTopLevelWindow(IntPtr hwnd)
    {
        // Ensure it's an Explorer process and a CabinetWClass top-level window.
        var info = _windowManager.GetWindowInfo(hwnd);
        if (info is null)
            return false;

        if (!string.Equals(NormalizeExeName(info.ProcessName), "explorer", StringComparison.OrdinalIgnoreCase))
            return false;

        return string.Equals(info.ClassName, ExplorerWindowClass, StringComparison.OrdinalIgnoreCase);
    }

    private async Task ConvertWindowToTab(
        IntPtr sourceTopLevel,
        IntPtr preferredTabHandle = default,
        string? preferredLocation = null,
        bool sourceHiddenAlready = false)
    {
        var sw = Stopwatch.StartNew();

        bool sourceHidden = sourceHiddenAlready;
        bool converted = false;

        try
        {
            string? location = preferredLocation;
            bool hasReadyLocation = IsRealFileSystemLocation(location);

            IntPtr sourceTabHandle = preferredTabHandle;
            if (!hasReadyLocation && (sourceTabHandle == IntPtr.Zero || !NativeMethods.IsWindow(sourceTabHandle)))
            {
                // Explorer may fire SHOW very early; wait briefly for tab child window(s).
                List<IntPtr> initialTabs = await WaitForTabHandles(sourceTopLevel, timeoutMs: 250);
                if (initialTabs.Count == 0)
                {
                    _logger.Info($"Skip convert: cannot find tab child for 0x{sourceTopLevel.ToInt64():X}");
                    return;
                }

                // Only fold single-tab windows.
                if (initialTabs.Count != 1)
                {
                    _logger.Info($"Skip convert: not a single-tab window 0x{sourceTopLevel.ToInt64():X}");
                    return;
                }

                sourceTabHandle = initialTabs[0];
            }

            if (!hasReadyLocation && sourceTabHandle == IntPtr.Zero)
            {
                _logger.Info($"Skip convert: cannot find active tab for 0x{sourceTopLevel.ToInt64():X}");
                return;
            }

            IntPtr targetTopLevel = PickTargetExplorerWindow(sourceTopLevel);
            if (targetTopLevel == IntPtr.Zero)
            {
                _logger.Info($"Skip convert: no target explorer window found for 0x{sourceTopLevel.ToInt64():X}");
                return;
            }

            if (targetTopLevel == sourceTopLevel)
                return;

            if (!sourceHidden)
                sourceHidden = _windowManager.Hide(sourceTopLevel);

            if (!hasReadyLocation)
            {
                location = await WaitForRealLocationByTabHandle(sourceTabHandle, sourceTopLevel, timeoutMs: 500);
                hasReadyLocation = IsRealFileSystemLocation(location);
            }

            if (string.IsNullOrWhiteSpace(location))
            {
                _logger.Info($"Skip convert: location not ready for 0x{sourceTopLevel.ToInt64():X}");
                return;
            }

            if (await TryActivateExistingTabByLocation(location, excludeTopLevel: sourceTopLevel))
            {
                bool closeExistingPosted = _windowManager.Close(sourceTopLevel);
                if (!closeExistingPosted && _windowManager.IsAlive(sourceTopLevel))
                {
                    _logger.Warn($"Convert warning: failed to close source window after existing-tab activation 0x{sourceTopLevel.ToInt64():X}");
                    return;
                }

                converted = true;
                _logger.Info($"Converted Explorer window to existing tab: {location} ({sw.ElapsedMilliseconds} ms)");
                return;
            }

            bool success = await OpenLocationInNewTab(targetTopLevel, location);
            if (!success)
            {
                _logger.Info($"Convert failed: open location as tab failed for {location}");
                return;
            }

            bool closePosted = _windowManager.Close(sourceTopLevel);
            if (!closePosted && _windowManager.IsAlive(sourceTopLevel))
            {
                _logger.Warn($"Convert warning: failed to close source window 0x{sourceTopLevel.ToInt64():X}");
                return;
            }

            converted = true;

            _logger.Info($"Converted Explorer window to tab: {location} ({sw.ElapsedMilliseconds} ms)");
        }
        finally
        {
            if (!converted && sourceHidden && _windowManager.IsAlive(sourceTopLevel))
                _windowManager.Show(sourceTopLevel);
        }
    }

    private static async Task<List<IntPtr>> WaitForTabHandles(IntPtr explorerTopLevel, int timeoutMs)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            List<IntPtr> tabs = GetAllTabHandles(explorerTopLevel);
            if (tabs.Count > 0)
                return tabs;
            await Task.Delay(20);
        }

        return [];
    }

    private IntPtr PickTargetExplorerWindow(IntPtr exclude)
    {
        // Prefer current foreground Explorer window if it isn't the excluded one.
        IntPtr foreground = NativeMethods.GetForegroundWindow();
        if (foreground != IntPtr.Zero && foreground != exclude && IsUsableTargetExplorerWindow(foreground))
            return foreground;

        // Otherwise prefer the last active Explorer window to preserve user's context.
        IntPtr lastForeground = _lastExplorerForeground;
        if (lastForeground != IntPtr.Zero && lastForeground != exclude && IsUsableTargetExplorerWindow(lastForeground))
            return lastForeground;

        // Otherwise pick a visible, non-minimized Explorer window with most tabs.
        // Do not fall back to hidden windows, which makes conversion look like a no-op.
        var candidates = EnumerateExplorerTopLevelWindows(includeInvisible: false)
            .Where(h => h != exclude && IsUsableTargetExplorerWindow(h))
            .ToList();

        if (candidates.Count == 0)
            return IntPtr.Zero;

        return candidates
            .OrderByDescending(CountTabs)
            .FirstOrDefault();
    }

    private bool IsUsableTargetExplorerWindow(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd))
            return false;

        if (!NativeMethods.IsWindowVisible(hwnd))
            return false;

        var placement = new NativeStructs.WINDOWPLACEMENT
        {
            length = (uint)Marshal.SizeOf<NativeStructs.WINDOWPLACEMENT>()
        };

        if (NativeMethods.GetWindowPlacement(hwnd, ref placement) &&
            placement.showCmd == (uint)NativeConstants.SW_SHOWMINIMIZED)
        {
            return false;
        }

        return IsExplorerTopLevelWindow(hwnd);
    }

    private IEnumerable<IntPtr> EnumerateExplorerTopLevelWindows(bool includeInvisible)
    {
        var result = new List<IntPtr>();

        NativeMethods.EnumWindows((hWnd, _) =>
        {
            if (!includeInvisible && !NativeMethods.IsWindowVisible(hWnd))
                return true;

            var classBuilder = new StringBuilder(256);
            NativeMethods.GetClassName(hWnd, classBuilder, classBuilder.Capacity);
            if (!string.Equals(classBuilder.ToString(), ExplorerWindowClass, StringComparison.OrdinalIgnoreCase))
                return true;

            var info = _windowManager.GetWindowInfo(hWnd);
            if (info is null)
                return true;

            if (!string.Equals(NormalizeExeName(info.ProcessName), "explorer", StringComparison.OrdinalIgnoreCase))
                return true;

            result.Add(hWnd);
            return true;
        }, IntPtr.Zero);

        return result;
    }

    private static int CountTabs(IntPtr explorerTopLevel)
    {
        int count = 0;
        IntPtr h = IntPtr.Zero;
        while (true)
        {
            h = NativeMethods.FindWindowEx(explorerTopLevel, h, ExplorerTabClass, null);
            if (h == IntPtr.Zero)
                break;
            count++;
        }
        return count;
    }

    private static IntPtr GetActiveTabHandle(IntPtr explorerTopLevel)
    {
        // Active tab is the first ShellTabWindowClass.
        return NativeMethods.FindWindowEx(explorerTopLevel, IntPtr.Zero, ExplorerTabClass, null);
    }

    private async Task<bool> OpenLocationInNewTab(IntPtr targetTopLevel, string location)
    {
        lock (_sendLock)
        {
            // Ensure target is foreground for the WM_COMMAND to land correctly.
            NativeMethods.ShowWindow(targetTopLevel, NativeConstants.SW_RESTORE);
            NativeMethods.BringWindowToTop(targetTopLevel);
            NativeMethods.SetForegroundWindow(targetTopLevel);
        }

        IntPtr activeTab = GetActiveTabHandle(targetTopLevel);
        if (activeTab == IntPtr.Zero)
            return false;

        IntPtr oldActiveTab = activeTab;
        List<IntPtr> before = GetAllTabHandles(targetTopLevel);

        NativeMethods.PostMessage(
            activeTab,
            (uint)NativeConstants.WM_COMMAND,
            new IntPtr(NativeConstants.EXPLORER_CMD_OPEN_NEW_TAB),
            IntPtr.Zero);

        IntPtr newTabHandle = await WaitForActiveTabChange(targetTopLevel, oldActiveTab, timeoutMs: 240);
        if (newTabHandle == IntPtr.Zero)
            newTabHandle = await WaitForNewTabHandle(targetTopLevel, before, timeoutMs: 240);

        if (newTabHandle == IntPtr.Zero)
        {
            _logger.Debug($"OpenLocationInNewTab: new tab handle not found for 0x{targetTopLevel.ToInt64():X}");
            return false;
        }

        // Prefer COM navigation (doesn't rely on focus / SendKeys).
        bool navigated = await NavigateTabByHandleWithRetry(newTabHandle, location, timeoutMs: 420);
        if (!navigated)
        {
            bool sendKeysOk = await NavigateViaAddressBar(location);
            if (!sendKeysOk)
            {
                CloseTab(newTabHandle);
                return false;
            }
        }

        TryBringToForeground(targetTopLevel);
        return true;
    }

    private static void TryBringToForeground(IntPtr topLevel)
    {
        if (topLevel == IntPtr.Zero || !NativeMethods.IsWindow(topLevel))
            return;

        NativeMethods.ShowWindow(topLevel, NativeConstants.SW_RESTORE);
        NativeMethods.BringWindowToTop(topLevel);

        if (NativeMethods.SetForegroundWindow(topLevel))
            return;

        IntPtr foreground = NativeMethods.GetForegroundWindow();
        if (foreground == IntPtr.Zero)
            return;

        uint foregroundThread = NativeMethods.GetWindowThreadProcessId(foreground, out _);
        uint targetThread = NativeMethods.GetWindowThreadProcessId(topLevel, out _);
        uint currentThread = NativeMethods.GetCurrentThreadId();

        bool currentAttached = false;
        bool targetAttached = false;

        try
        {
            if (foregroundThread != 0 && foregroundThread != currentThread)
            {
                currentAttached = NativeMethods.AttachThreadInput(currentThread, foregroundThread, true);
            }

            if (foregroundThread != 0 && targetThread != 0 && targetThread != foregroundThread)
            {
                targetAttached = NativeMethods.AttachThreadInput(targetThread, foregroundThread, true);
            }

            NativeMethods.BringWindowToTop(topLevel);
            NativeMethods.SetForegroundWindow(topLevel);
        }
        finally
        {
            if (targetAttached)
                NativeMethods.AttachThreadInput(targetThread, foregroundThread, false);

            if (currentAttached)
                NativeMethods.AttachThreadInput(currentThread, foregroundThread, false);
        }
    }

    private static async Task<string?> WaitForRealLocationByTabHandle(IntPtr tabHandle, IntPtr topLevelHwnd, int timeoutMs)
    {
        int start = Environment.TickCount;

        while (Environment.TickCount - start < timeoutMs)
        {
            string? location = await UiAsync(() =>
                TryGetLocationByTabHandleUi(tabHandle) ?? TryGetLocationByTopLevelUi(topLevelHwnd));

            if (IsRealFileSystemLocation(location))
                return location;

            await Task.Delay(40);
        }

        return null;
    }

    private static string? TryGetLocationByTopLevelUi(IntPtr topLevelHwnd)
    {
        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                dynamic win = tab;
                IntPtr hwnd = new IntPtr((int)win.HWND);
                if (hwnd != topLevelHwnd)
                    continue;

                return TryGetComLocation(tab);
            }
            catch
            {
                // ignore
            }
            finally
            {
                Marshal.FinalReleaseComObject(tab);
            }
        }

        return null;
    }

    private static string? TryGetLocationByTabHandleUi(IntPtr tabHandle)
    {
        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                if (GetTabHandle(tab) != tabHandle)
                    continue;

                return TryGetComLocation(tab);
            }
            catch
            {
                // ignore
            }
            finally
            {
                Marshal.FinalReleaseComObject(tab);
            }
        }

        return null;
    }

    private static async Task<bool> NavigateTabByHandleWithRetry(IntPtr tabHandle, string location, int timeoutMs)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            bool ok = await UiAsync(() => TryNavigateTabByHandleUi(tabHandle, location));
            if (ok)
                return true;
            await Task.Delay(80);
        }
        return false;
    }

    private static bool TryNavigateTabByHandleUi(IntPtr tabHandle, string location)
    {
        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                if (GetTabHandle(tab) != tabHandle)
                    continue;

                return TryNavigateComTab(tab, location);
            }
            catch
            {
                return false;
            }
            finally
            {
                Marshal.FinalReleaseComObject(tab);
            }
        }

        return false;
    }

    private static async Task<bool> WaitUntilTabLocationMatches(IntPtr tabHandle, string location, int timeoutMs, int pollMs)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            string? current = await UiAsync(() => TryGetLocationByTabHandleUi(tabHandle));
            if (!string.IsNullOrWhiteSpace(current) && PathsEquivalent(current, location))
                return true;
            await Task.Delay(pollMs);
        }
        return false;
    }

    private static async Task<IntPtr> WaitForActiveTabChange(IntPtr explorerTopLevel, IntPtr oldActiveTab, int timeoutMs)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            IntPtr current = GetActiveTabHandle(explorerTopLevel);
            if (current != IntPtr.Zero && current != oldActiveTab)
                return current;

            await Task.Delay(30);
        }

        return IntPtr.Zero;
    }

    private static Task<bool> NavigateViaAddressBar(string location)
    {
        // Clipboard + Ctrl+L is the most reliable without COM tab object binding.
        return System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
        {
            try
            {
                string original = ClipboardManager.GetClipboardText();
                ClipboardManager.SetClipboardText(location);

                SendKeys.SendWait("^{l}");
                Thread.Sleep(40);
                SendKeys.SendWait("^{v}");
                SendKeys.SendWait("{ENTER}");

                // Best-effort restore.
                Thread.Sleep(20);
                ClipboardManager.SetClipboardText(original);

                return true;
            }
            catch
            {
                return false;
            }
        }).Task;
    }

    private static List<IntPtr> GetAllTabHandles(IntPtr explorerTopLevel)
    {
        var result = new List<IntPtr>();
        IntPtr h = IntPtr.Zero;
        while (true)
        {
            h = NativeMethods.FindWindowEx(explorerTopLevel, h, ExplorerTabClass, null);
            if (h == IntPtr.Zero)
                break;
            result.Add(h);
        }
        return result;
    }

    private static async Task<IntPtr> WaitForNewTabHandle(IntPtr explorerTopLevel, List<IntPtr> before, int timeoutMs)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            foreach (IntPtr tab in GetAllTabHandles(explorerTopLevel))
            {
                if (!before.Contains(tab))
                    return tab;
            }
            await Task.Delay(35);
        }
        return IntPtr.Zero;
    }

    private static void CloseTab(IntPtr tabHandle)
    {
        if (tabHandle == IntPtr.Zero)
            return;

        NativeMethods.PostMessage(
            tabHandle,
            (uint)NativeConstants.WM_COMMAND,
            new IntPtr(NativeConstants.EXPLORER_CMD_CLOSE_TAB),
            new IntPtr(1));
    }

    private static List<object> GetShellWindowsSnapshotUi()
    {
        var result = new List<object>();

        object? windows = GetShellWindowsCollectionUi();
        if (windows is null)
            return result;

        try
        {
            dynamic dynWindows = windows!;
            int count = (int)dynWindows.Count;

            for (int i = 0; i < count; i++)
            {
                object? window = null;
                try
                {
                    window = dynWindows.Item(i);
                    if (window is null)
                        continue;

                    // Filter to explorer tabs.
                    string fullName = (string?)((dynamic)window).FullName ?? string.Empty;
                    if (!fullName.EndsWith(ExplorerExe, StringComparison.OrdinalIgnoreCase))
                    {
                        Marshal.FinalReleaseComObject(window);
                        continue;
                    }

                    result.Add(window);
                }
                catch
                {
                    if (window is not null)
                        Marshal.FinalReleaseComObject(window);
                }
            }
        }
        finally
        {
            // Intentionally keep the cached ShellWindows collection alive.
        }

        return result;
    }

    private static object? GetShellWindowsCollectionUi()
    {
        // Must run on UI(STA) thread.
        lock (ShellWindowsInitLock)
        {
            try
            {
                int tid = Environment.CurrentManagedThreadId;
                if (_shellWindows is not null && _shellWindowsThreadId == tid)
                    return _shellWindows;

                // If called on a different thread, recreate to match COM apartment.
                if (_shellWindows is not null)
                {
                    Marshal.FinalReleaseComObject(_shellWindows);
                    _shellWindows = null;
                }

                Type? shellWindowsType = Type.GetTypeFromCLSID(ShellWindowsClsid);
                if (shellWindowsType is null)
                    return null;

                _shellWindows = Activator.CreateInstance(shellWindowsType);
                if (_shellWindows is null)
                    return null;

                _shellWindowsThreadId = tid;
                return _shellWindows;
            }
            catch
            {
                return null;
            }
        }
    }

    private static IntPtr GetTabHandle(object comTab)
    {
        try
        {
            if (comTab is not WinTab.App.ExplorerTabUtilityPort.Interop.IServiceProvider sp)
                return IntPtr.Zero;

            Guid guid = ShellBrowserGuid;
            Guid iid = ShellBrowserGuid;
            int hr = sp.QueryService(ref guid, ref iid, out IShellBrowser? shellBrowser);
            if (hr != 0 || shellBrowser is null)
                return IntPtr.Zero;

            try
            {
                _ = shellBrowser.GetWindow(out nint handle);
                return new IntPtr(handle);
            }
            finally
            {
                Marshal.FinalReleaseComObject(shellBrowser);
            }
        }
        catch
        {
            return IntPtr.Zero;
        }
    }

    private static async Task<string?> WaitForRealLocation(object comTab, int timeoutMs)
    {
        int start = Environment.TickCount;
        string? last = null;
        int stable = 0;

        while (Environment.TickCount - start < timeoutMs)
        {
            string? location = TryGetComLocation(comTab);
            if (IsRealFileSystemLocation(location))
            {
                if (string.Equals(location, last, StringComparison.OrdinalIgnoreCase))
                    stable++;
                else
                {
                    last = location;
                    stable = 1;
                }

                if (stable >= 2)
                    return location;
            }

            await Task.Delay(70);
        }

        return null;
    }

    private static bool TryGetComTabLocation(object? newTabObj, IntPtr tabHandle, out string? location)
    {
        location = null;

        if (newTabObj is not null)
        {
            location = TryGetComLocation(newTabObj);
            return !string.IsNullOrWhiteSpace(location);
        }

        // Fallback: re-enumerate and match by tab handle.
        try
        {
            Type? shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType is null)
                return false;

            dynamic shell = Activator.CreateInstance(shellType)!;
            dynamic windows = shell.Windows();
            int count = (int)windows.Count;
            for (int i = 0; i < count; i++)
            {
                object? item = null;
                try
                {
                    item = windows.Item(i);
                    if (item is null)
                        continue;

                    string fullName = (string?)((dynamic)item).FullName ?? string.Empty;
                    if (!fullName.EndsWith(ExplorerExe, StringComparison.OrdinalIgnoreCase))
                    {
                        Marshal.FinalReleaseComObject(item);
                        continue;
                    }

                    if (GetTabHandle(item) != tabHandle)
                    {
                        Marshal.FinalReleaseComObject(item);
                        continue;
                    }

                    location = TryGetComLocation(item);
                    Marshal.FinalReleaseComObject(item);
                    return !string.IsNullOrWhiteSpace(location);
                }
                catch
                {
                    if (item is not null)
                        Marshal.FinalReleaseComObject(item);
                }
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private static string? TryGetComLocation(object comTab)
    {
        try
        {
            dynamic win = comTab;
            string url = (string?)win.LocationURL ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(url) && Uri.TryCreate(url, UriKind.Absolute, out Uri? uri) && uri.IsFile)
                return Uri.UnescapeDataString(uri.LocalPath);
        }
        catch
        {
            // ignore
        }

        try
        {
            dynamic win = comTab;
            string path = (string?)win.Document?.Folder?.Self?.Path ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(path))
                return path;
        }
        catch
        {
            // ignore
        }

        return null;
    }

    private static bool TryNavigateComTab(object comTab, string location)
    {
        try
        {
            dynamic win = comTab;

            if (location.Contains('#'))
            {
                Type? shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType is null)
                    return false;

                dynamic shell = Activator.CreateInstance(shellType)!;
                object? folder = null;
                try
                {
                    folder = shell.NameSpace(location);
                    if (folder is null)
                        return false;
                    win.Navigate2(folder);
                    return true;
                }
                finally
                {
                    if (folder is not null)
                        Marshal.FinalReleaseComObject(folder);
                    Marshal.FinalReleaseComObject(shell);
                }
            }

            win.Navigate2(location);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsRealFileSystemLocation(string? location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return false;

        if (location.StartsWith("shell::", StringComparison.OrdinalIgnoreCase) ||
            location.StartsWith("::", StringComparison.OrdinalIgnoreCase))
            return false;

        if (location.StartsWith("\\\\", StringComparison.OrdinalIgnoreCase))
            return true;

        if (location.Length >= 3 && char.IsLetter(location[0]) && location[1] == ':' && (location[2] == '\\' || location[2] == '/'))
            return true;

        return false;
    }

    private static bool PathsEquivalent(string a, string b)
    {
        try
        {
            string left = NormalizePathForCompare(Path.GetFullPath(a));
            string right = NormalizePathForCompare(Path.GetFullPath(b));
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(
                NormalizePathForCompare(a),
                NormalizePathForCompare(b),
                StringComparison.OrdinalIgnoreCase);
        }
    }

    private static bool IsChildPathOf(string parentPath, string childPath)
    {
        try
        {
            string parent = NormalizePathForCompare(Path.GetFullPath(parentPath));
            string child = NormalizePathForCompare(Path.GetFullPath(childPath));

            if (child.Length <= parent.Length)
                return false;

            if (!child.StartsWith(parent, StringComparison.OrdinalIgnoreCase))
                return false;

            return child[parent.Length] == '\\';
        }
        catch
        {
            return false;
        }
    }

    private static string NormalizePathForCompare(string path)
    {
        if (string.IsNullOrEmpty(path))
            return string.Empty;

        return path
            .Replace('/', '\\')
            .TrimEnd('\\')
            .Normalize(NormalizationForm.FormKC);
    }

    private static string NormalizeExeName(string processName)
    {
        string normalized = processName.Trim();
        if (normalized.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            normalized = normalized[..^4];
        return normalized;
    }

    [ComImport]
    [Guid("B196B284-BAB4-101A-B69C-00AA00341D07")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IConnectionPointContainerNative
    {
        void EnumConnectionPoints(out IntPtr ppEnum);
        void FindConnectionPoint(ref Guid riid, out IConnectionPointNative ppCP);
    }

    [ComImport]
    [Guid("B196B286-BAB4-101A-B69C-00AA00341D07")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IConnectionPointNative
    {
        void GetConnectionInterface(out Guid pIID);
        void GetConnectionPointContainer([MarshalAs(UnmanagedType.Interface)] out IConnectionPointContainerNative ppCPC);
        void Advise([MarshalAs(UnmanagedType.IUnknown)] object pUnkSink, out int pdwCookie);
        void Unadvise(int dwCookie);
        void EnumConnections(out IntPtr ppEnum);
    }

    [ComVisible(true)]
    [Guid("FE4106E0-399A-11D0-A48C-00A0C90A8F39")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDShellWindowsEvents
    {
        [DispId(200)]
        void WindowRegistered(int lCookie);

        [DispId(201)]
        void WindowRevoked(int lCookie);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class ShellWindowsEventsSink : IDShellWindowsEvents
    {
        private readonly Action<int> _onWindowRegistered;

        public ShellWindowsEventsSink(Action<int> onWindowRegistered)
        {
            _onWindowRegistered = onWindowRegistered;
        }

        public void WindowRegistered(int lCookie)
        {
            _onWindowRegistered(lCookie);
        }

        public void WindowRevoked(int lCookie)
        {
            // not used
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        if (_createEventHook != IntPtr.Zero)
            NativeMethods.UnhookWinEvent(_createEventHook);

        if (_shellWindowsConnectionPoint is not null && _shellWindowsConnectionCookie != 0)
        {
            try
            {
                _shellWindowsConnectionPoint.Unadvise(_shellWindowsConnectionCookie);
            }
            catch
            {
                // ignore
            }
        }

        if (_shellWindowsConnectionPoint is not null)
        {
            Marshal.FinalReleaseComObject(_shellWindowsConnectionPoint);
            _shellWindowsConnectionPoint = null;
        }

        _shellWindowsEventsSink = null;
        _shellWindowsConnectionCookie = 0;

        _windowEvents.WindowShown -= OnWindowShown;
        _windowEvents.WindowDestroyed -= OnWindowDestroyed;
        _windowEvents.WindowForegroundChanged -= OnWindowForegroundChanged;
    }
}
