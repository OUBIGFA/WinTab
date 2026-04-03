// Adapted from ExplorerTabUtility (MIT License).
// Source: E:\_BIGFA Free\_code\ExplorerTabUtility

using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Threading;
using WinTab.App.ExplorerTabUtilityPort.Interop;
using WinTab.App.Services;
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
public sealed class ExplorerTabHookService : IDisposable, IExplorerAutoConvertController
{
    public async Task<bool> OpenLocationAsTabAsync(string location, IntPtr clickTimeForeground = default)
    {
        if (string.IsNullOrWhiteSpace(location))
            return false;

        location = location.Trim();
        OpenTargetInfo targetInfo = OpenTargetClassifier.Classify(location);
        if (targetInfo.RequiresNativeShellLaunch)
            return AppEnvironment.TryOpenTargetFallback(location, _logger);

        bool handledInExistingWindow = await TryOpenLocationInExistingExplorerContextAsync(location, clickTimeForeground);
        if (handledInExistingWindow)
            return true;

        _logger.Warn($"[AutoConvert] OpenLocationAsTabAsync: in-window open failed, fallback to normal explorer open for '{location}'.");
        return TryLaunchExplorerWindowWithBypassPolicy(location);
    }

    public async Task<bool> OpenInterceptedLocationAsTabAsync(string location, IntPtr clickTimeForeground = default)
    {
        if (string.IsNullOrWhiteSpace(location))
            return false;

        location = location.Trim();
        OpenTargetInfo targetInfo = OpenTargetClassifier.Classify(location);
        if (targetInfo.RequiresNativeShellLaunch)
            return AppEnvironment.TryOpenTargetFallback(location, _logger);

        try
        {
            await _interceptOpenGate.WaitAsync(_cts.Token);
        }
        catch (OperationCanceledException)
        {
            return false;
        }

        try
        {
            bool handled = await TryOpenLocationInActiveExplorerWindowOnlyAsync(location, clickTimeForeground);
            if (!handled)
            {
                _logger.Warn($"[Intercept] Open request cannot be handled in current active Explorer window: '{location}'.");
            }

            return handled;
        }
        finally
        {
            _interceptOpenGate.Release();
        }
    }

    private async Task<bool> TryOpenLocationInActiveExplorerWindowOnlyAsync(string location, IntPtr clickTimeForeground)
    {
        // Match native Explorer behavior first: when opening from inside Explorer,
        // navigate current active tab unless user explicitly enabled child-folder-in-new-tab.
        CurrentNavigateAttemptResult currentNavigate = await TryNavigateCurrentActiveTabLikeExplorer(location, clickTimeForeground);
        if (currentNavigate.NavigatedCurrentTab)
            return true;

        if (!ShouldContinueWithTabReuseAfterCurrentNavigateAttempt(
                currentNavigate.NavigatedCurrentTab,
                currentNavigate.RequiredCurrentWindowNavigation))
        {
            _logger.Warn($"[Intercept] Required current-window navigation failed, skipped tab-hijack fallback: {location}");

            // Best-effort safety fallback: keep shell-native behavior instead of forcing tab hijack.
            _nativeBrowseFallbackBypass.Register(location, NativeBrowseFallbackBypassTtl);
            bool launched = TryLaunchExplorerWindow(location);
            if (!launched)
            {
                _nativeBrowseFallbackBypass.Revoke(location);
                _logger.Warn($"[Intercept] Native fallback launch failed after current-window navigation failure: {location}");
            }

            return launched;
        }

        IntPtr targetTopLevel = PickInterceptTargetExplorerWindow(clickTimeForeground);
        if (targetTopLevel == IntPtr.Zero)
        {
            _logger.Info($"[Intercept] No active Explorer window available; fallback to standalone Explorer launch: {location}");
            bool launched = TryLaunchExplorerWindowWithBypassPolicy(location);
            if (!launched)
            {
                _logger.Warn($"[Intercept] Fallback Explorer launch failed: {location}");
            }

            return launched;
        }

        // Strict policy for intercepted requests:
        // - only operate in the CURRENT active Explorer window
        // - never activate another Explorer window just because it already has the same path
        if (await TryActivateExistingTabByLocation(location, requiredTopLevel: targetTopLevel))
            return true;

        for (int attempt = 0; attempt < 3; attempt++)
        {
            bool opened = await OpenLocationInNewTab(targetTopLevel, location);
            if (opened)
            {
                TryBringToForeground(targetTopLevel);
                _logger.Info($"[Intercept] Opened location in current active Explorer window tab: {location}");
                return true;
            }

            try
            {
                await Task.Delay(80, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                return false;
            }

            IntPtr refreshed = PickInterceptTargetExplorerWindow(clickTimeForeground);
            if (refreshed == IntPtr.Zero)
                break;

            targetTopLevel = refreshed;
        }

        return false;
    }

    private async Task<bool> TryOpenLocationInExistingExplorerContextAsync(string location, IntPtr clickTimeForeground)
    {
        // Keep tab reuse behavior: if same location already exists, activate it instead of creating a duplicate tab.
        if (await TryActivateExistingTabByLocation(location))
            return true;

        IntPtr targetTopLevel = PickInterceptTargetExplorerWindow(clickTimeForeground);
        if (targetTopLevel == IntPtr.Zero)
            return false;

        // Intercepted requests should open directly as a NEW TAB in the active Explorer context.
        for (int attempt = 0; attempt < 3; attempt++)
        {
            bool opened = await OpenLocationInNewTab(targetTopLevel, location);
            if (opened)
            {
                TryBringToForeground(targetTopLevel);
                _logger.Info($"[Intercept] Opened location in active Explorer new tab: {location}");
                return true;
            }

            try
            {
                await Task.Delay(80, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                return false;
            }

            IntPtr refreshed = PickInterceptTargetExplorerWindow(clickTimeForeground);
            if (refreshed == IntPtr.Zero)
                break;

            targetTopLevel = refreshed;
        }

        return false;
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

    private bool TryLaunchExplorerWindowWithBypassPolicy(string location)
    {
        bool shouldBypassAutoConvert = ShouldBypassAutoConvertForLocation(location);
        if (shouldBypassAutoConvert)
            _nativeBrowseFallbackBypass.Register(location, NativeBrowseFallbackBypassTtl);

        bool launched = TryLaunchExplorerWindow(location);
        if (!launched && shouldBypassAutoConvert)
            _nativeBrowseFallbackBypass.Revoke(location);

        return launched;
    }

    private async Task<CurrentNavigateAttemptResult> TryNavigateCurrentActiveTabLikeExplorer(string location, IntPtr clickTimeForeground)
    {
        // clickTimeForeground is captured by the handler process at the exact moment the user clicked,
        // before the shell hands off control. This is the authoritative foreground for determining
        // if the open action originated from within an Explorer window.
        IntPtr explorerWindow = IntPtr.Zero;
        if (clickTimeForeground != IntPtr.Zero && IsExplorerTopLevelWindow(clickTimeForeground))
            explorerWindow = clickTimeForeground;

        // Fallback: if the foreground wasn't captured (legacy pipe protocol), check now.
        if (explorerWindow == IntPtr.Zero)
        {
            IntPtr fg = NativeMethods.GetForegroundWindow();
            if (IsExplorerTopLevelWindow(fg))
                explorerWindow = fg;
        }

        if (explorerWindow == IntPtr.Zero || !NativeMethods.IsWindow(explorerWindow))
            return CurrentNavigateAttemptResult.NotRequiredAndNotNavigated();

        bool requiredCurrentWindowNavigation = !_settings.OpenChildFolderInNewTabFromActiveTab;

        IntPtr activeTab = GetActiveTabHandle(explorerWindow);
        if (activeTab == IntPtr.Zero || !NativeMethods.IsWindow(activeTab))
            return requiredCurrentWindowNavigation
                ? CurrentNavigateAttemptResult.RequiredButNotNavigated()
                : CurrentNavigateAttemptResult.NotRequiredAndNotNavigated();

        bool forceNavigateCurrentForSameParent = false;
        if (_settings.OpenNewTabFromActiveTabPath && IsRealFileSystemLocation(location))
        {
            string? currentLocation = await UiAsync(() => TryGetLocationByTabHandleUi(activeTab));
            if (IsRealFileSystemLocation(currentLocation))
            {
                string? currentParent = TryGetNormalizedParentPath(currentLocation!);
                string? targetParent = TryGetNormalizedParentPath(location);
                if (currentParent is not null && targetParent is not null)
                {
                    forceNavigateCurrentForSameParent = string.Equals(currentParent, targetParent, StringComparison.OrdinalIgnoreCase);
                }
            }
        }

        if (forceNavigateCurrentForSameParent)
            requiredCurrentWindowNavigation = true;

        if (!requiredCurrentWindowNavigation)
            return CurrentNavigateAttemptResult.NotRequiredAndNotNavigated();

        if (forceNavigateCurrentForSameParent)
        {
            _logger.Info($"[Intercept] Same-parent browse detected with inherit-current-path enabled; keep native current-tab navigation: {location}");
        }

        // Use a tight timeout for current-tab in-place navigation: COM navigate is
        // fire-and-commit. When TryNavigateTabByHandleUi returns true the instruction
        // is accepted; we only retry briefly in case the COM object is momentarily
        // unavailable. No confirmation wait — that would add 60-600 ms of polling
        // with no correctness benefit.
        bool navigated = await NavigateTabByHandleWithRetry(activeTab, location, timeoutMs: CurrentTabNavigateTimeoutMs, ct: _cts.Token);
        if (!navigated)
        {
            // Keep native browse behavior in the current Explorer window even when COM navigate is flaky.
            TryBringToForeground(explorerWindow);
            bool navigatedViaAddressBar = await NavigateViaAddressBar(location);
            if (!navigatedViaAddressBar)
                return CurrentNavigateAttemptResult.RequiredButNotNavigated();

            SuppressNewTabDefaultBehavior(explorerWindow, AddressBarFallbackSuppressionTtl);
            RememberNewTabDefaultSnapshot(explorerWindow, CountTabs(explorerWindow), location);
            _logger.Info($"Navigated active Explorer tab via address bar fallback: {location}");
            return CurrentNavigateAttemptResult.RequiredAndNavigated();
        }

        TryBringToForeground(explorerWindow);
        RememberNewTabDefaultSnapshot(explorerWindow, CountTabs(explorerWindow), location);
        _logger.Info($"Navigated active Explorer tab directly: {location}");
        return CurrentNavigateAttemptResult.RequiredAndNavigated();
    }

    private static bool ShouldContinueWithTabReuseAfterCurrentNavigateAttempt(bool navigatedCurrentTab, bool requiredCurrentWindowNavigation)
    {
        return !navigatedCurrentTab && !requiredCurrentWindowNavigation;
    }

    private static readonly Guid ShellBrowserGuid = typeof(IShellBrowser).GUID;
    private static readonly Guid ShellWindowsClsid = new("9BA05972-F6A8-11CF-A442-00A0C90A8F39");
    private static readonly Guid ShellWindowsEventsGuid = new("FE4106E0-399A-11D0-A48C-00A0C90A8F39");
    private const string ExplorerExe = "explorer.exe";
    private const string ExplorerTabClass = "ShellTabWindowClass";
    private const string ExplorerWindowClass = "CabinetWClass";

    // Timeout constants for current-tab in-place navigation.
    // COM navigate is fire-and-commit: when TryNavigateTabByHandleUi returns true the
    // navigation instruction is accepted by Explorer's COM interface. We only retry
    // briefly in case the COM object is momentarily unavailable.
    internal const int CurrentTabNavigateTimeoutMs = 200;
    internal const int CurrentTabNavigateRetryMs = 100;
    private static readonly TimeSpan NativeBrowseFallbackBypassTtl = TimeSpan.FromSeconds(4);
    private static readonly TimeSpan AddressBarFallbackSuppressionTtl = TimeSpan.FromMilliseconds(1200);
    private static readonly TimeSpan ExternalSuppressionDefaultTtl = TimeSpan.FromMilliseconds(1200);

    private static readonly object ShellWindowsInitLock = new();
    private static object? _shellWindows;
    private static int _shellWindowsThreadId;
    private static int _clipboardOperationInProgress;
    private static readonly ConcurrentDictionary<IntPtr, DateTimeOffset> _externalSuppressNewTabDefaultUntil = new();

    private readonly IWindowEventSource _windowEvents;
    private readonly IWindowManager _windowManager;
    private readonly ShellLocationIdentityService _locationIdentity;
    private readonly NativeBrowseFallbackBypassStore _nativeBrowseFallbackBypass;
    private readonly AppSettings _settings;
    private readonly Logger _logger;
    private volatile bool _autoConvertEnabled;

    private readonly Action<int> _windowRegisteredHandler;
    private readonly NativeMethods.WinEventDelegate _createEventCallback;
    private IConnectionPointNative? _shellWindowsConnectionPoint;
    private object? _shellWindowsEventsSink;
    private int _shellWindowsConnectionCookie;
    private bool _shellWindowRegisteredHooked;
    private readonly WinEventHookManager _winEventHookManager;

    private readonly ConcurrentDictionary<IntPtr, DateTimeOffset> _pending = new();
    private readonly ConcurrentDictionary<IntPtr, DateTimeOffset> _suppressNewTabDefaultUntil = new();
    private readonly ConcurrentDictionary<IntPtr, byte> _knownExplorerTopLevels = new();
    private readonly ConcurrentDictionary<IntPtr, byte> _earlyHiddenExplorer = new();
    private readonly ConcurrentDictionary<IntPtr, ConcurrentDictionary<IntPtr, int>> _tabIndexCache = new();
    private readonly ConcurrentDictionary<IntPtr, byte> _newTabDefaultInFlight = new();
    private readonly ConcurrentDictionary<IntPtr, NewTabDefaultSnapshot> _newTabDefaultSnapshots = new();
    private readonly ConcurrentDictionary<int, byte> _windowRegisteredInFlight = new();
    private readonly CancellationTokenSource _cts = new();
    private readonly ConcurrentDictionary<uint, bool> _explorerPidCache = new();
    private readonly object _sendLock = new();
    private readonly SemaphoreSlim _interceptOpenGate = new(1, 1);
    private System.Windows.Threading.DispatcherTimer? _cleanupTimer;

    private IntPtr _lastExplorerForeground;
    private IntPtr _mainExplorerTopLevel;

    private bool _disposed;

    private readonly ShellComNavigator _shellComNavigator;

    public ExplorerTabHookService(
        IWindowEventSource windowEvents,
        IWindowManager windowManager,
        ShellLocationIdentityService locationIdentity,
        AppSettings settings,
        Logger logger)
    {
        _windowEvents = windowEvents;
        _windowManager = windowManager;
        _locationIdentity = locationIdentity;
        _nativeBrowseFallbackBypass = new NativeBrowseFallbackBypassStore(locationIdentity);
        _settings = settings;
        _logger = logger;
        _windowRegisteredHandler = cookie => OnShellWindowRegisteredSafe(cookie);
        _createEventCallback = OnExplorerObjectCreate;
        _winEventHookManager = new WinEventHookManager(logger, _createEventCallback);
        _shellComNavigator = new ShellComNavigator(logger);

        _autoConvertEnabled = settings.EnableAutoConvertExplorerWindows;

        _windowEvents.WindowForegroundChanged += OnWindowForegroundChanged;
        _windowEvents.WindowDestroyed += OnWindowDestroyed;

        _windowEvents.WindowShown += OnWindowShown;

        IntPtr currentForeground = NativeMethods.GetForegroundWindow();
        if (currentForeground != IntPtr.Zero && IsExplorerTopLevelWindow(currentForeground))
            _lastExplorerForeground = currentForeground;

        // Seed cache with existing Explorer windows so we don't treat normal
        // "show" events (navigation, layout changes) as new windows.
        foreach (IntPtr h in EnumerateExplorerTopLevelWindows(includeInvisible: true))
        {
            _knownExplorerTopLevels.TryAdd(h, 0);
            SeedNewTabDefaultSnapshot(h);
        }

        _mainExplorerTopLevel = PickEtuTargetExplorerWindow(exclude: IntPtr.Zero);

        _shellWindowRegisteredHooked = TryHookShellWindowRegistered();
        if (_shellWindowRegisteredHooked)
        {
            _logger.Info("ExplorerTabHookService: ShellWindows.WindowRegistered hooked.");
        }
        else if (_autoConvertEnabled)
        {
            _logger.Warn("ExplorerTabHookService: ShellWindows.WindowRegistered hook unavailable, using WindowShown fallback.");
        }

        // Install EVENT_OBJECT_CREATE hook via dedicated manager:
        // - auto-convert mode: reduce flash by early hiding new Explorer windows
        // - on-demand mode: observe newly created tab windows for default new-tab behavior
        _winEventHookManager.Start();

        if (_autoConvertEnabled)
            _logger.Info("ExplorerTabHookService started (ETU-style, no COMReference).");
        else
            _logger.Info("ExplorerTabHookService started in on-demand mode (auto-convert disabled).");

        // Periodic cleanup of stale dictionary entries (every 5 minutes).
        var dispatcher = TryGetUiDispatcher();
        if (dispatcher is not null)
        {
            _cleanupTimer = new System.Windows.Threading.DispatcherTimer(
                TimeSpan.FromMinutes(5),
                System.Windows.Threading.DispatcherPriority.Background,
                (_, _) => CleanupStaleDictionaryEntries(),
                dispatcher);
            _cleanupTimer.Start();
        }
    }

    private readonly record struct CurrentNavigateAttemptResult(bool NavigatedCurrentTab, bool RequiredCurrentWindowNavigation)
    {
        public static CurrentNavigateAttemptResult NotRequiredAndNotNavigated() => new(false, false);
        public static CurrentNavigateAttemptResult RequiredAndNavigated() => new(true, true);
        public static CurrentNavigateAttemptResult RequiredButNotNavigated() => new(false, true);
    }

    private readonly record struct NewTabDefaultSnapshot(int TabCount, string? ActiveLocation, DateTimeOffset CapturedAtUtc);
    private readonly record struct RegisteredCandidate(IntPtr TopLevelHwnd, IntPtr TabHandle, string? Location, bool IsKnownTopLevel);
    private readonly record struct ExistingTabCandidate(IntPtr TopLevelHwnd, IntPtr TabHandle, string Location);

    private async Task<bool> TryActivateExistingTabByLocation(string location, IntPtr excludeTopLevel = default, IntPtr requiredTopLevel = default)
    {
        ExistingTabCandidate? candidate = await UiAsync(() => TryFindExistingTabByLocationUi(location, excludeTopLevel, requiredTopLevel));
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

    private ExistingTabCandidate? TryFindExistingTabByLocationUi(string location, IntPtr excludeTopLevel, IntPtr requiredTopLevel)
    {
        if (!IsTabNavigableLocation(location))
            return null;

        IntPtr foreground = NativeMethods.GetForegroundWindow();
        IntPtr lastForeground = _lastExplorerForeground;

        ExistingTabCandidate? firstMatch = null;
        ExistingTabCandidate? preferredLastForegroundMatch = null;

        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                string? tabLocation = _shellComNavigator.TryGetComLocation(tab);
                if (tabLocation is null ||
                    !IsTabNavigableLocation(tabLocation) ||
                    !_locationIdentity.AreEquivalent(tabLocation, location))
                    continue;

                IntPtr tabHandle = GetTabHandle(tab);
                IntPtr topLevel = IntPtr.Zero;

                if (tabHandle == IntPtr.Zero)
                {
                    dynamic win = tab;
                    topLevel = new IntPtr((int)win.HWND);
                    if (topLevel == IntPtr.Zero || !NativeMethods.IsWindow(topLevel))
                        continue;
                    tabHandle = topLevel;
                }
                else
                {
                    if (!NativeMethods.IsWindow(tabHandle))
                        continue;
                    topLevel = NativeMethods.GetAncestor(tabHandle, NativeConstants.GA_ROOT);
                    if (topLevel == IntPtr.Zero || !NativeMethods.IsWindow(topLevel))
                        continue;
                }

                if (excludeTopLevel != IntPtr.Zero && topLevel == excludeTopLevel)
                    continue;

                if (requiredTopLevel != IntPtr.Zero && topLevel != requiredTopLevel)
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

        return await WaitForActiveTabHandle(explorerTopLevel, tabHandle, timeoutMs: timeoutMs, ct: _cts.Token);
    }

    private async Task<int> ProbeTabIndexForHandle(IntPtr explorerTopLevel, IntPtr tabHandle, List<IntPtr> tabs, int skipPrimary, int skipSecondary)
    {
        for (int i = 0; i < tabs.Count; i++)
        {
            if (i == skipPrimary || i == skipSecondary)
                continue;

            if (!TrySelectTabByIndex(explorerTopLevel, i))
                continue;

            if (await WaitForActiveTabHandle(explorerTopLevel, tabHandle, timeoutMs: 120, ct: _cts.Token))
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
            IntPtr sent = NativeMethods.SendMessageTimeout(
                explorerTopLevel,
                (uint)NativeConstants.WM_SETREDRAW,
                IntPtr.Zero,
                IntPtr.Zero,
                NativeConstants.SMTO_ABORTIFHUNG,
                200,
                out _);

            redrawLocked = sent != IntPtr.Zero;
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
            NativeMethods.SendMessageTimeout(
                explorerTopLevel,
                (uint)NativeConstants.WM_SETREDRAW,
                new IntPtr(1),
                IntPtr.Zero,
                NativeConstants.SMTO_ABORTIFHUNG,
                200,
                out _);

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

    private static async Task<bool> WaitForActiveTabHandle(IntPtr explorerTopLevel, IntPtr expectedTabHandle, int timeoutMs, CancellationToken ct = default)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            if (GetActiveTabHandle(explorerTopLevel) == expectedTabHandle)
                return true;

            await Task.Delay(25, ct);
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

    private void OnShellWindowRegisteredSafe(int cookie)
    {
        if (_disposed)
            return;

        if (!_windowRegisteredInFlight.TryAdd(cookie, 0))
        {
            _logger.Debug($"ShellWindowRegistered: duplicate cookie skipped ({cookie}).");
            return;
        }

        // COM callback must never throw — wrap in Task.Run with full exception isolation.
        _ = Task.Run(async () =>
        {
            try
            {
                if (_disposed)
                    return;

                await OnShellWindowRegisteredCore(cookie);
            }
            catch (Exception ex)
            {
                _logger.Debug($"ShellWindowRegistered handler failed (cookie={cookie}): {ex.Message}");
            }
            finally
            {
                _windowRegisteredInFlight.TryRemove(cookie, out _);
            }
        }, _cts.Token);
    }

    private async Task OnShellWindowRegisteredCore(int cookie)
    {
        if (_disposed)
            return;

        try
        {
            bool shouldHandleNewTabDefault = _settings.OpenNewTabFromActiveTabPath;

            RegisteredCandidate? candidate = await WaitForRegisteredCandidate(
                cookie,
                timeoutMs: 1200,
                includeKnownTopLevel: shouldHandleNewTabDefault);

            if (candidate is null)
                return;

            if (shouldHandleNewTabDefault && candidate.Value.IsKnownTopLevel)
            {
                if (_cts.IsCancellationRequested)
                    return;

                _ = Task.Run(async () =>
                {
                    try
                    {
                        if (_disposed)
                            return;

                        await TryApplyNewTabDefaultLocation(
                            candidate.Value.TopLevelHwnd,
                            candidate.Value.TabHandle,
                            candidate.Value.Location);
                    }
                    catch (OperationCanceledException)
                    {
                        // ignore
                    }
                    catch (Exception ex)
                    {
                        _logger.Debug($"Failed to apply new-tab default location (cookie={cookie}, hwnd=0x{candidate.Value.TopLevelHwnd.ToInt64():X}): {ex.Message}");
                    }
                }, _cts.Token);
            }

            if (!candidate.Value.IsKnownTopLevel &&
                !ShouldConvertWindowLocationToTab(candidate.Value.Location))
            {
                RestoreEarlyHiddenWindow(candidate.Value.TopLevelHwnd);
                _logger.Info($"Skip convert: keep native shell window for {candidate.Value.Location}");
                return;
            }

            if (!_autoConvertEnabled || candidate.Value.IsKnownTopLevel)
                return;

            IntPtr hwnd = candidate.Value.TopLevelHwnd;
            if (!_pending.TryAdd(hwnd, DateTimeOffset.UtcNow))
                return;

            _logger.Info($"Explorer candidate window registered: 0x{hwnd.ToInt64():X}");

            bool hiddenImmediately = _earlyHiddenExplorer.TryRemove(hwnd, out _);
            if (!hiddenImmediately)
                hiddenImmediately = await TryHideExplorerWindowAggressively(hwnd, timeoutMs: 180);

            if (_cts.IsCancellationRequested)
            {
                _pending.TryRemove(hwnd, out _);
                return;
            }

            _ = Task.Run(async () =>
            {
                try
                {
                    if (_disposed)
                        return;

                    await ConvertWindowToTab(
                        hwnd,
                        candidate.Value.TabHandle,
                        candidate.Value.Location,
                        sourceHiddenAlready: hiddenImmediately);
                }
                catch (OperationCanceledException)
                {
                    // ignore
                }
                catch (Exception ex)
                {
                    _logger.Error($"Explorer tab hook failed for 0x{hwnd.ToInt64():X} (cookie={cookie}).", ex);
                }
                finally
                {
                    _pending.TryRemove(hwnd, out _);
                }
            }, _cts.Token);
        }
        catch (Exception ex)
        {
            _logger.Debug($"ShellWindowRegistered handler error: {ex.Message}");
        }
    }

    private async Task TryApplyNewTabDefaultLocation(IntPtr topLevel, IntPtr suggestedTabHandle, string? suggestedLocation)
    {
        if (_disposed || !_settings.OpenNewTabFromActiveTabPath)
            return;

        if (topLevel == IntPtr.Zero || !NativeMethods.IsWindow(topLevel) || !IsExplorerTopLevelWindow(topLevel))
            return;

        try
        {
            if (IsNewTabDefaultSuppressed(topLevel))
                return;

            if (!_newTabDefaultSnapshots.TryGetValue(topLevel, out NewTabDefaultSnapshot baseline))
                return;

            await Task.Delay(80, _cts.Token);

            int currentTabCount = CountTabs(topLevel);
            if (!ShouldApplyCachedNewTabAlignment(baseline.TabCount, currentTabCount, baseline.ActiveLocation))
                return;

            string sourceLocation = baseline.ActiveLocation!;

            IntPtr newTabHandle = GetActiveTabHandle(topLevel);
            if (newTabHandle == IntPtr.Zero || !NativeMethods.IsWindow(newTabHandle))
                newTabHandle = suggestedTabHandle;

            if (newTabHandle == IntPtr.Zero || !NativeMethods.IsWindow(newTabHandle))
                return;

            string? newTabLocation = await UiAsync(() => TryGetLocationByTabHandleUi(newTabHandle));
            if (!IsRealFileSystemLocation(newTabLocation) && newTabHandle == suggestedTabHandle)
                newTabLocation = suggestedLocation;

            // Sometimes the registration callback arrives before Explorer actually
            // switches active tab; retry once before deciding alignment behavior.
            if (IsRealFileSystemLocation(newTabLocation))
            {
                await Task.Delay(120, _cts.Token);

                IntPtr refreshedActiveTab = GetActiveTabHandle(topLevel);
                if (refreshedActiveTab != IntPtr.Zero && NativeMethods.IsWindow(refreshedActiveTab))
                    newTabHandle = refreshedActiveTab;

                newTabLocation = await UiAsync(() => TryGetLocationByTabHandleUi(newTabHandle));
                if (!IsRealFileSystemLocation(newTabLocation) && newTabHandle == suggestedTabHandle)
                    newTabLocation = suggestedLocation;
            }

            if (ShouldSkipNewTabAlignment(newTabLocation, sourceLocation, _locationIdentity))
                return;

            bool navigated = await NavigateTabByHandleWithRetry(newTabHandle, sourceLocation, timeoutMs: 420, ct: _cts.Token);
            if (!navigated)
                return;

            _ = await WaitUntilTabLocationMatches(newTabHandle, sourceLocation, timeoutMs: 420, pollMs: 60);
            _logger.Info($"Aligned new Explorer tab to active tab location: {sourceLocation}");
            RememberNewTabDefaultSnapshot(topLevel, currentTabCount, sourceLocation);
        }
        finally
        {
            try
            {
                await RefreshNewTabDefaultSnapshot(topLevel);
            }
            catch (OperationCanceledException)
            {
                // ignore
            }
            catch
            {
                // ignore
            }
        }
    }

    private static bool ShouldSkipNewTabAlignment(
        string? newTabLocation,
        string? sourceLocation,
        ShellLocationIdentityService locationIdentity)
    {
        ArgumentNullException.ThrowIfNull(locationIdentity);

        if (!IsRealFileSystemLocation(sourceLocation))
            return true;

        if (!IsRealFileSystemLocation(newTabLocation))
            return false;

        return locationIdentity.AreEquivalent(newTabLocation, sourceLocation);
    }

    private static bool ShouldApplyCachedNewTabAlignment(int previousTabCount, int currentTabCount, string? cachedSourceLocation)
    {
        if (previousTabCount <= 0)
            return false;

        if (currentTabCount <= previousTabCount)
            return false;

        return IsPhysicalFileSystemLocation(cachedSourceLocation);
    }

    private void SeedNewTabDefaultSnapshot(IntPtr topLevel)
    {
        if (topLevel == IntPtr.Zero || !NativeMethods.IsWindow(topLevel) || !IsExplorerTopLevelWindow(topLevel))
            return;

        try
        {
            NewTabDefaultSnapshot snapshot = UiAsync(() => CaptureNewTabDefaultSnapshotUi(topLevel)).GetAwaiter().GetResult();
            _newTabDefaultSnapshots[topLevel] = snapshot;
        }
        catch
        {
            // ignore
        }
    }

    private async Task RefreshNewTabDefaultSnapshot(IntPtr topLevel)
    {
        if (_disposed || topLevel == IntPtr.Zero || !NativeMethods.IsWindow(topLevel) || !IsExplorerTopLevelWindow(topLevel))
            return;

        NewTabDefaultSnapshot snapshot = await UiAsync(() => CaptureNewTabDefaultSnapshotUi(topLevel));
        _newTabDefaultSnapshots[topLevel] = snapshot;
    }

    private NewTabDefaultSnapshot CaptureNewTabDefaultSnapshotUi(IntPtr topLevel)
    {
        int tabCount = CountTabs(topLevel);
        string? activeLocation = null;

        IntPtr activeTab = GetActiveTabHandle(topLevel);
        if (activeTab != IntPtr.Zero && NativeMethods.IsWindow(activeTab))
            activeLocation = TryGetLocationByTabHandleUi(activeTab);

        return new NewTabDefaultSnapshot(tabCount, activeLocation, DateTimeOffset.UtcNow);
    }

    private void RememberNewTabDefaultSnapshot(IntPtr topLevel, int tabCount, string? activeLocation)
    {
        if (topLevel == IntPtr.Zero)
            return;

        _newTabDefaultSnapshots[topLevel] = new NewTabDefaultSnapshot(tabCount, activeLocation, DateTimeOffset.UtcNow);
    }

    public static void SuppressNewTabDefaultForWindow(IntPtr explorerTopLevel, TimeSpan? duration = null)
    {
        if (explorerTopLevel == IntPtr.Zero)
            return;

        TimeSpan effectiveDuration = duration ?? ExternalSuppressionDefaultTtl;
        if (effectiveDuration <= TimeSpan.Zero)
            return;

        _externalSuppressNewTabDefaultUntil[explorerTopLevel] = DateTimeOffset.UtcNow.Add(effectiveDuration);
    }

    private void SuppressNewTabDefaultBehavior(IntPtr explorerTopLevel, TimeSpan duration)
    {
        if (explorerTopLevel == IntPtr.Zero)
            return;

        _suppressNewTabDefaultUntil[explorerTopLevel] = DateTimeOffset.UtcNow.Add(duration);
    }

    private bool IsNewTabDefaultSuppressed(IntPtr explorerTopLevel)
    {
        DateTimeOffset now = DateTimeOffset.UtcNow;

        bool localSuppressed = false;
        if (_suppressNewTabDefaultUntil.TryGetValue(explorerTopLevel, out DateTimeOffset untilUtc))
        {
            if (untilUtc > now)
                localSuppressed = true;
            else
                _suppressNewTabDefaultUntil.TryRemove(explorerTopLevel, out _);
        }

        bool externalSuppressed = false;
        if (_externalSuppressNewTabDefaultUntil.TryGetValue(explorerTopLevel, out DateTimeOffset externalUntilUtc))
        {
            if (externalUntilUtc > now)
                externalSuppressed = true;
            else
                _externalSuppressNewTabDefaultUntil.TryRemove(explorerTopLevel, out _);
        }

        return localSuppressed || externalSuppressed;
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

        var classBuilder = new StringBuilder(32);
        NativeMethods.GetClassName(hwnd, classBuilder, classBuilder.Capacity);
        string windowClass = classBuilder.ToString();

        if (string.Equals(windowClass, ExplorerTabClass, StringComparison.OrdinalIgnoreCase))
        {
            if (!_settings.OpenNewTabFromActiveTabPath)
                return;

            IntPtr topLevel = NativeMethods.GetAncestor(hwnd, NativeConstants.GA_ROOT);
            if (topLevel == IntPtr.Zero || !NativeMethods.IsWindow(topLevel))
                return;

            // Throttle: skip if already processing a new-tab-default for this window.
            if (!_newTabDefaultInFlight.TryAdd(topLevel, 0))
                return;

            if (_cts.IsCancellationRequested)
            {
                _newTabDefaultInFlight.TryRemove(topLevel, out _);
                return;
            }

            _ = Task.Run(async () =>
            {
                try
                {
                    if (_disposed)
                        return;

                    await TryApplyNewTabDefaultLocation(topLevel, hwnd, suggestedLocation: null);
                }
                catch (OperationCanceledException)
                {
                    // ignore
                }
                catch (Exception ex)
                {
                    _logger.Debug($"Failed to handle tab-create event (top=0x{topLevel.ToInt64():X}, tab=0x{hwnd.ToInt64():X}): {ex.Message}");
                }
                finally
                {
                    _newTabDefaultInFlight.TryRemove(topLevel, out _);
                }
            }, _cts.Token);

            return;
        }

        if (!string.Equals(windowClass, ExplorerWindowClass, StringComparison.OrdinalIgnoreCase))
            return;

        if (!_autoConvertEnabled)
            return;

        if (_pending.ContainsKey(hwnd) || _knownExplorerTopLevels.ContainsKey(hwnd))
            return;

        // DWM-cloak first (instant, prevents the window from ever being
        // composited) then also SW_HIDE as a safety net.  Together these
        // eliminate the flash that a standalone ShowWindow(SW_HIDE) cannot
        // prevent due to the inherent race with Explorer's first paint.
        // Track in _earlyHiddenExplorer whenever we successfully altered
        // visibility by either mechanism, so the restore paths can find it.
        bool suppressed = _windowManager.SuppressVisibility(hwnd);
        bool hidden = _windowManager.Hide(hwnd);
        if (suppressed || hidden)
            _earlyHiddenExplorer.TryAdd(hwnd, 0);
    }

    private async Task<RegisteredCandidate?> WaitForRegisteredCandidate(int cookie, int timeoutMs, bool includeKnownTopLevel)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            RegisteredCandidate? candidate = await UiAsync(() =>
                TryTakeRegisteredCandidateByCookieUi(cookie, includeKnownTopLevel));

            if (candidate is not null)
                return candidate;

            if (_autoConvertEnabled)
                candidate = await UiAsync(() => TryTakeRegisteredCandidateUi(includeKnownTopLevel: false));

            if (candidate is not null)
                return candidate;

            await Task.Delay(25, _cts.Token);
        }

        return null;
    }

    private RegisteredCandidate? TryTakeRegisteredCandidateByCookieUi(int cookie, bool includeKnownTopLevel)
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

            bool knownTopLevel = _knownExplorerTopLevels.ContainsKey(topLevel);
            if (knownTopLevel && !includeKnownTopLevel)
                return null;

            if (!IsExplorerTopLevelWindow(topLevel))
                return null;

            if (!knownTopLevel)
                _knownExplorerTopLevels.TryAdd(topLevel, 0);

            IntPtr tabHandle = GetTabHandle(tab);
            string? location = _shellComNavigator.TryGetComLocation(tab);

            return new RegisteredCandidate(topLevel, tabHandle, location, knownTopLevel);
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

    private RegisteredCandidate? TryTakeRegisteredCandidateUi(bool includeKnownTopLevel)
    {
        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                dynamic win = tab;
                IntPtr topLevel = new IntPtr((int)win.HWND);
                if (topLevel == IntPtr.Zero)
                    continue;

                bool knownTopLevel = _knownExplorerTopLevels.ContainsKey(topLevel);
                if (knownTopLevel && !includeKnownTopLevel)
                    continue;

                if (!IsExplorerTopLevelWindow(topLevel))
                    continue;

                if (!knownTopLevel)
                    _knownExplorerTopLevels.TryAdd(topLevel, 0);

                IntPtr tabHandle = GetTabHandle(tab);
                string? location = _shellComNavigator.TryGetComLocation(tab);

                return new RegisteredCandidate(topLevel, tabHandle, location, knownTopLevel);
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
        {
            _lastExplorerForeground = hwnd;
            _mainExplorerTopLevel = hwnd;

            if (!_cts.IsCancellationRequested)
            {
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await RefreshNewTabDefaultSnapshot(hwnd);
                    }
                    catch (OperationCanceledException)
                    {
                        // ignore
                    }
                    catch
                    {
                        // ignore
                    }
                }, _cts.Token);
            }
        }
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

        if (_mainExplorerTopLevel == hwnd)
            _mainExplorerTopLevel = IntPtr.Zero;

        _knownExplorerTopLevels.TryRemove(hwnd, out _);
        _earlyHiddenExplorer.TryRemove(hwnd, out _);
        _pending.TryRemove(hwnd, out _);
        _newTabDefaultInFlight.TryRemove(hwnd, out _);
        _suppressNewTabDefaultUntil.TryRemove(hwnd, out _);
        _newTabDefaultSnapshots.TryRemove(hwnd, out _);
        _externalSuppressNewTabDefaultUntil.TryRemove(hwnd, out _);
        _tabIndexCache.TryRemove(hwnd, out _);
    }

    private void OnWindowShown(object? sender, IntPtr hwnd)
    {
        if (_disposed || hwnd == IntPtr.Zero || !_autoConvertEnabled)
            return;

        // Even when WindowRegistered is hooked, SHOW can still arrive earlier or be used for re-hide.
        TryQueueShownCandidate(hwnd);
    }

    private void TryQueueShownCandidate(IntPtr hwnd)
    {
        if (!_autoConvertEnabled)
            return;

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

        if (_cts.IsCancellationRequested)
        {
            _pending.TryRemove(hwnd, out _);
            return;
        }

        _ = Task.Run(async () =>
        {
            try
            {
                if (_disposed)
                    return;

                bool hiddenImmediately = await TryHideExplorerWindowAggressively(hwnd, timeoutMs: 180);
                await ConvertWindowToTab(hwnd, sourceHiddenAlready: hiddenImmediately);
            }
            catch (OperationCanceledException)
            {
                // ignore
            }
            catch (Exception ex)
            {
                _logger.Error($"Explorer tab hook failed for 0x{hwnd.ToInt64():X} (source=WindowShown).", ex);
            }
            finally
            {
                _pending.TryRemove(hwnd, out _);
            }
        }, _cts.Token);
    }

    private async Task<bool> TryHideExplorerWindowAggressively(IntPtr hwnd, int timeoutMs)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd))
            return false;

        // DWM-cloak immediately — takes effect before the next DWM frame,
        // eliminating any visual flash regardless of subsequent polling.
        _windowManager.SuppressVisibility(hwnd);

        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            _windowManager.SuppressVisibility(hwnd);
            if (_windowManager.Hide(hwnd) && !NativeMethods.IsWindowVisible(hwnd))
                return true;

            try
            {
                await Task.Delay(15, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                return false;
            }

            if (!NativeMethods.IsWindow(hwnd))
                return false;
        }

        _windowManager.SuppressVisibility(hwnd);
        return _windowManager.Hide(hwnd);
    }

    private bool IsExplorerTopLevelWindow(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd))
            return false;

        // Fast class name check first (cheap, no cross-process call for process info).
        var classBuilder = new StringBuilder(64);
        NativeMethods.GetClassName(hwnd, classBuilder, classBuilder.Capacity);
        if (!string.Equals(classBuilder.ToString(), ExplorerWindowClass, StringComparison.OrdinalIgnoreCase))
            return false;

        // Check PID against cached explorer PIDs.
        NativeMethods.GetWindowThreadProcessId(hwnd, out uint pid);
        if (pid == 0)
            return false;

        if (_explorerPidCache.TryGetValue(pid, out bool isExplorer))
            return isExplorer;

        // Cap cache size to prevent unbounded growth.
        if (_explorerPidCache.Count > 50)
            _explorerPidCache.Clear();

        // Cache miss: do the expensive check once per PID.
        isExplorer = false;
        try
        {
            using var process = Process.GetProcessById((int)pid);
            isExplorer = string.Equals(process.ProcessName, "explorer", StringComparison.OrdinalIgnoreCase);
        }
        catch (ArgumentException) { }
        catch (InvalidOperationException) { }

        _explorerPidCache[pid] = isExplorer;
        return isExplorer;
    }

    private async Task ConvertWindowToTab(
        IntPtr sourceTopLevel,
        IntPtr preferredTabHandle = default,
        string? preferredLocation = null,
        bool sourceHiddenAlready = false)
    {
        var sw = Stopwatch.StartNew();

        bool sourceHidden = sourceHiddenAlready;
        bool visibilitySuppressionStarted = false;
        bool converted = false;
        using var visibilitySuppressionCts = new CancellationTokenSource();
        Task? visibilitySuppressionTask = null;

        try
        {
            string? location = preferredLocation;
            bool hasReadyLocation = IsTabNavigableLocation(location);
            if (hasReadyLocation && !ShouldConvertWindowLocationToTab(location))
            {
                _logger.Info($"Skip convert: keep native shell window for {location}");
                return;
            }

            IntPtr sourceTabHandle = preferredTabHandle;
            if (!hasReadyLocation && (sourceTabHandle == IntPtr.Zero || !NativeMethods.IsWindow(sourceTabHandle)))
            {
                // Explorer may fire SHOW very early; wait briefly for tab child window(s).
                List<IntPtr> initialTabs = await WaitForTabHandles(sourceTopLevel, timeoutMs: 250, ct: _cts.Token);
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

            IntPtr targetTopLevel = PickEtuTargetExplorerWindow(sourceTopLevel);
            if (targetTopLevel == IntPtr.Zero)
            {
                _logger.Info($"Skip convert: no target explorer window found for 0x{sourceTopLevel.ToInt64():X}");
                return;
            }

            if (targetTopLevel == sourceTopLevel)
                return;

            if (hasReadyLocation && location is not null && _nativeBrowseFallbackBypass.TryConsume(location))
            {
                _logger.Info($"Skip convert: native-browse fallback bypass matched for {location}");
                return;
            }

            if (!sourceHidden)
            {
                _windowManager.SuppressVisibility(sourceTopLevel);
                sourceHidden = _windowManager.Hide(sourceTopLevel);
            }

            visibilitySuppressionTask = KeepWindowHiddenUntilConversionCompletes(sourceTopLevel, visibilitySuppressionCts.Token);
            visibilitySuppressionStarted = true;

            if (!hasReadyLocation)
            {
                location = await WaitForRealLocationByTabHandle(sourceTabHandle, sourceTopLevel, timeoutMs: 500, ct: _cts.Token);
                hasReadyLocation = IsTabNavigableLocation(location);
            }

            if (string.IsNullOrWhiteSpace(location))
            {
                _logger.Info($"Skip convert: location not ready for 0x{sourceTopLevel.ToInt64():X}");
                return;
            }

            if (!ShouldConvertWindowLocationToTab(location))
            {
                _logger.Info($"Skip convert: keep native shell window for {location}");
                return;
            }

            if (_nativeBrowseFallbackBypass.TryConsume(location))
            {
                _logger.Info($"Skip convert: native-browse fallback bypass matched for {location}");
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
            visibilitySuppressionCts.Cancel();

            if (visibilitySuppressionTask is not null)
            {
                try
                {
                    await visibilitySuppressionTask;
                }
                catch (OperationCanceledException)
                {
                    // expected on shutdown
                }
            }

            if (!converted &&
                (sourceHidden || visibilitySuppressionStarted) &&
                _windowManager.IsAlive(sourceTopLevel))
            {
                _windowManager.RestoreVisibility(sourceTopLevel);
                _windowManager.Show(sourceTopLevel);
            }
        }
    }

    private async Task KeepWindowHiddenUntilConversionCompletes(IntPtr hwnd, CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd))
                return;

            // Re-cloak every iteration (idempotent, cheap) so DWM never
            // composites the window even if Explorer toggles WS_VISIBLE.
            _windowManager.SuppressVisibility(hwnd);

            if (NativeMethods.IsWindowVisible(hwnd))
                _windowManager.Hide(hwnd);

            await Task.Delay(30, ct);
        }
    }

    private static async Task<List<IntPtr>> WaitForTabHandles(IntPtr explorerTopLevel, int timeoutMs, CancellationToken ct = default)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            List<IntPtr> tabs = GetAllTabHandles(explorerTopLevel);
            if (tabs.Count > 0)
                return tabs;
            await Task.Delay(20, ct);
        }

        return [];
    }

    private IntPtr PickInterceptTargetExplorerWindow(IntPtr clickTimeForeground)
    {
        // Intercepted-open policy:
        // Prefer active-context Explorer first, but keep robust fallback to an existing Explorer window
        // so requests do not fail when click-time foreground handle is stale or non-Explorer.
        if (clickTimeForeground != IntPtr.Zero)
        {
            IntPtr clickRoot = NativeMethods.GetAncestor(clickTimeForeground, NativeConstants.GA_ROOT);
            if (clickRoot != IntPtr.Zero && IsUsableTargetExplorerWindow(clickRoot, allowMinimized: true))
                return clickRoot;
        }

        IntPtr currentForeground = NativeMethods.GetForegroundWindow();
        if (currentForeground != IntPtr.Zero)
        {
            IntPtr currentRoot = NativeMethods.GetAncestor(currentForeground, NativeConstants.GA_ROOT);
            if (currentRoot != IntPtr.Zero && IsUsableTargetExplorerWindow(currentRoot, allowMinimized: true))
                return currentRoot;

            if (IsUsableTargetExplorerWindow(currentForeground, allowMinimized: true))
                return currentForeground;
        }

        if (_lastExplorerForeground != IntPtr.Zero && IsUsableTargetExplorerWindow(_lastExplorerForeground, allowMinimized: true))
            return _lastExplorerForeground;

        return PickTargetExplorerWindow(exclude: IntPtr.Zero, allowMinimizedFallback: true);
    }

    private IntPtr PickTargetExplorerWindow(IntPtr exclude, bool allowMinimizedFallback = false)
    {
        IntPtr foreground = NativeMethods.GetForegroundWindow();
        IntPtr lastForeground = _lastExplorerForeground;
        List<(IntPtr Hwnd, bool IsMinimized, int TabCount)> candidates = CollectReusableExplorerCandidates(exclude, allowMinimizedFallback);
        return SelectBestExplorerTargetCandidate(candidates, foreground, lastForeground, allowMinimizedFallback);
    }

    private IntPtr PickEtuTargetExplorerWindow(IntPtr exclude)
    {
        // ETU-style third-party takeover: keep a stable "main" Explorer window
        // and fold newly opened standalone windows into that window's tabs.
        IntPtr main = _mainExplorerTopLevel;
        if (main != IntPtr.Zero && main != exclude && IsUsableTargetExplorerWindow(main))
            return main;

        List<(IntPtr Hwnd, bool IsMinimized, int TabCount)> candidates = CollectReusableExplorerCandidates(exclude, allowMinimizedFallback: true);
        candidates.Reverse();

        IntPtr picked = SelectBestExplorerTargetCandidate(
            candidates,
            preferredForeground: main,
            preferredLastForeground: _lastExplorerForeground,
            allowMinimizedFallback: true);
        if (picked == IntPtr.Zero)
            return IntPtr.Zero;

        _mainExplorerTopLevel = picked;
        return picked;
    }

    private List<(IntPtr Hwnd, bool IsMinimized, int TabCount)> CollectReusableExplorerCandidates(IntPtr exclude, bool allowMinimizedFallback)
    {
        var windows = new List<(IntPtr Hwnd, bool IsVisible, bool IsMinimized, int TabCount)>();

        foreach (IntPtr hwnd in EnumerateExplorerTopLevelWindows(includeInvisible: allowMinimizedFallback))
        {
            if (hwnd == IntPtr.Zero || hwnd == exclude || !NativeMethods.IsWindow(hwnd))
                continue;

            bool isMinimized = IsWindowMinimized(hwnd);
            windows.Add((hwnd, NativeMethods.IsWindowVisible(hwnd), isMinimized, CountTabs(hwnd)));
        }

        return BuildReusableExplorerCandidates(windows, exclude, allowMinimizedFallback);
    }

    private static List<(IntPtr Hwnd, bool IsMinimized, int TabCount)> BuildReusableExplorerCandidates(
        IReadOnlyList<(IntPtr Hwnd, bool IsVisible, bool IsMinimized, int TabCount)> windows,
        IntPtr exclude,
        bool allowMinimizedFallback)
    {
        var candidates = new List<(IntPtr Hwnd, bool IsMinimized, int TabCount)>();

        foreach ((IntPtr Hwnd, bool IsVisible, bool IsMinimized, int TabCount) window in windows)
        {
            if (window.Hwnd == IntPtr.Zero || window.Hwnd == exclude)
                continue;

            bool keepInvisibleMinimized = allowMinimizedFallback && window.IsMinimized;
            if (!window.IsVisible && !keepInvisibleMinimized)
                continue;

            candidates.Add((window.Hwnd, window.IsMinimized, window.TabCount));
        }

        return candidates;
    }

    private static IntPtr SelectBestExplorerTargetCandidate(
        IReadOnlyList<(IntPtr Hwnd, bool IsMinimized, int TabCount)> candidates,
        IntPtr preferredForeground,
        IntPtr preferredLastForeground,
        bool allowMinimizedFallback)
    {
        if (candidates.Count == 0)
            return IntPtr.Zero;

        static bool MatchesPreferred(
            (IntPtr Hwnd, bool IsMinimized, int TabCount) candidate,
            IntPtr preferredHwnd,
            bool allowMinimized)
        {
            if (preferredHwnd == IntPtr.Zero || candidate.Hwnd != preferredHwnd)
                return false;

            return allowMinimized || !candidate.IsMinimized;
        }

        foreach ((IntPtr Hwnd, bool IsMinimized, int TabCount) candidate in candidates)
        {
            if (MatchesPreferred(candidate, preferredForeground, allowMinimized: false))
                return candidate.Hwnd;
        }

        foreach ((IntPtr Hwnd, bool IsMinimized, int TabCount) candidate in candidates)
        {
            if (MatchesPreferred(candidate, preferredLastForeground, allowMinimized: false))
                return candidate.Hwnd;
        }

        IntPtr bestNonMinimized = IntPtr.Zero;
        int bestNonMinimizedTabs = -1;
        foreach ((IntPtr Hwnd, bool IsMinimized, int TabCount) candidate in candidates)
        {
            if (candidate.IsMinimized)
                continue;

            if (candidate.TabCount > bestNonMinimizedTabs)
            {
                bestNonMinimizedTabs = candidate.TabCount;
                bestNonMinimized = candidate.Hwnd;
            }
        }

        if (bestNonMinimized != IntPtr.Zero || !allowMinimizedFallback)
            return bestNonMinimized;

        foreach ((IntPtr Hwnd, bool IsMinimized, int TabCount) candidate in candidates)
        {
            if (MatchesPreferred(candidate, preferredForeground, allowMinimized: true))
                return candidate.Hwnd;
        }

        foreach ((IntPtr Hwnd, bool IsMinimized, int TabCount) candidate in candidates)
        {
            if (MatchesPreferred(candidate, preferredLastForeground, allowMinimized: true))
                return candidate.Hwnd;
        }

        IntPtr bestAny = IntPtr.Zero;
        int bestAnyTabs = -1;
        foreach ((IntPtr Hwnd, bool IsMinimized, int TabCount) candidate in candidates)
        {
            if (candidate.TabCount > bestAnyTabs)
            {
                bestAnyTabs = candidate.TabCount;
                bestAny = candidate.Hwnd;
            }
        }

        return bestAny;
    }

    private bool IsUsableTargetExplorerWindow(IntPtr hwnd, bool allowMinimized = false)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd))
            return false;

        if (!NativeMethods.IsWindowVisible(hwnd))
            return false;

        if (!allowMinimized && IsWindowMinimized(hwnd))
            return false;

        return IsExplorerTopLevelWindow(hwnd);
    }

    private static bool IsWindowMinimized(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd))
            return false;

        var placement = new NativeStructs.WINDOWPLACEMENT
        {
            length = (uint)Marshal.SizeOf<NativeStructs.WINDOWPLACEMENT>()
        };

        return NativeMethods.GetWindowPlacement(hwnd, ref placement) &&
               placement.showCmd == (uint)NativeConstants.SW_SHOWMINIMIZED;
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

        SuppressNewTabDefaultBehavior(targetTopLevel, TimeSpan.FromMilliseconds(1500));

        NativeMethods.PostMessage(
            activeTab,
            (uint)NativeConstants.WM_COMMAND,
            new IntPtr(NativeConstants.EXPLORER_CMD_OPEN_NEW_TAB),
            IntPtr.Zero);

        IntPtr newTabHandle = await WaitForActiveTabChange(targetTopLevel, oldActiveTab, timeoutMs: 240, ct: _cts.Token);
        if (newTabHandle == IntPtr.Zero)
            newTabHandle = await WaitForNewTabHandle(targetTopLevel, before, timeoutMs: 240, ct: _cts.Token);

        if (newTabHandle == IntPtr.Zero)
        {
            _logger.Debug($"OpenLocationInNewTab: new tab handle not found for 0x{targetTopLevel.ToInt64():X}");
            return false;
        }

        // Prefer COM navigation (doesn't rely on focus / SendKeys).
        bool navigated = await NavigateTabByHandleWithRetry(newTabHandle, location, timeoutMs: 420, ct: _cts.Token);
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
        RememberNewTabDefaultSnapshot(targetTopLevel, CountTabs(targetTopLevel), location);
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

    private async Task<string?> WaitForRealLocationByTabHandle(IntPtr tabHandle, IntPtr topLevelHwnd, int timeoutMs, CancellationToken ct = default)
    {
        int start = Environment.TickCount;

        while (Environment.TickCount - start < timeoutMs)
        {
            string? location = await UiAsync(() =>
                TryGetLocationByTabHandleUi(tabHandle) ?? TryGetLocationByTopLevelUi(topLevelHwnd));

            if (IsTabNavigableLocation(location))
                return location;

            await Task.Delay(60, ct);
        }

        return null;
    }

    private string? TryGetLocationByTopLevelUi(IntPtr topLevelHwnd)
    {
        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                dynamic win = tab;
                IntPtr hwnd = new IntPtr((int)win.HWND);
                if (hwnd != topLevelHwnd)
                    continue;

                return _shellComNavigator.TryGetComLocation(tab);
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

    private string? TryGetLocationByTabHandleUi(IntPtr tabHandle)
    {
        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                IntPtr th = GetTabHandle(tab);
                if (th == IntPtr.Zero)
                {
                    dynamic win = tab;
                    th = new IntPtr((int)win.HWND);
                }

                if (th != tabHandle)
                    continue;

                return _shellComNavigator.TryGetComLocation(tab);
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

    private async Task<bool> NavigateTabByHandleWithRetry(IntPtr tabHandle, string location, int timeoutMs, CancellationToken ct = default)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            bool ok = await UiAsync(() => TryNavigateTabByHandleUi(tabHandle, location));
            if (ok)
                return true;
            await Task.Delay(CurrentTabNavigateRetryMs, ct);
        }
        return false;
    }

    private bool TryNavigateTabByHandleUi(IntPtr tabHandle, string location)
    {
        foreach (object tab in GetShellWindowsSnapshotUi())
        {
            try
            {
                IntPtr th = GetTabHandle(tab);
                if (th == IntPtr.Zero)
                {
                    dynamic win = tab;
                    th = new IntPtr((int)win.HWND);
                }

                if (th != tabHandle)
                    continue;

                return _shellComNavigator.TryNavigateComTab(tab, location);
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

    private async Task<bool> WaitUntilTabLocationMatches(IntPtr tabHandle, string location, int timeoutMs, int pollMs)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            string? current = await UiAsync(() => TryGetLocationByTabHandleUi(tabHandle));
            if (!string.IsNullOrWhiteSpace(current) && _locationIdentity.AreEquivalent(current, location))
                return true;
            await Task.Delay(pollMs);
        }
        return false;
    }

    private static async Task<IntPtr> WaitForActiveTabChange(IntPtr explorerTopLevel, IntPtr oldActiveTab, int timeoutMs, CancellationToken ct = default)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            IntPtr current = GetActiveTabHandle(explorerTopLevel);
            if (current != IntPtr.Zero && current != oldActiveTab)
                return current;

            await Task.Delay(30, ct);
        }

        return IntPtr.Zero;
    }

    private static async Task<bool> NavigateViaAddressBar(string location)
    {
        // Safe-guard against extremely long paths crashing the UI automation (SendKeys).
        if (string.IsNullOrEmpty(location) || location.Length > 2048)
            return false;

        // Clipboard + Ctrl+L is the most reliable without COM tab object binding.
        // Use async delays instead of Thread.Sleep to avoid blocking the UI thread.
        var dispatcher = System.Windows.Application.Current.Dispatcher;
        if (dispatcher is null)
            return false;

        // Re-entrancy guard: prevent concurrent clipboard manipulation.
        if (!TryEnterClipboardOperation())
            return false;

        try
        {
            string original = await dispatcher.InvokeAsync(() =>
            {
                string orig = ClipboardManager.GetClipboardText();
                ClipboardManager.SetClipboardText(location);
                SendKeys.SendWait("^{l}");
                return orig;
            }).Task;

            await Task.Delay(40);

            await dispatcher.InvokeAsync(() =>
            {
                SendKeys.SendWait("^{v}");
                SendKeys.SendWait("{ENTER}");
            }).Task;

            await Task.Delay(20);

            await dispatcher.InvokeAsync(() =>
            {
                ClipboardManager.SetClipboardText(original);
            }).Task;

            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            ExitClipboardOperation();
        }
    }

    private static bool TryEnterClipboardOperation()
    {
        return Interlocked.CompareExchange(ref _clipboardOperationInProgress, 1, 0) == 0;
    }

    private static void ExitClipboardOperation()
    {
        Volatile.Write(ref _clipboardOperationInProgress, 0);
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

    private static async Task<IntPtr> WaitForNewTabHandle(IntPtr explorerTopLevel, List<IntPtr> before, int timeoutMs, CancellationToken ct = default)
    {
        int start = Environment.TickCount;
        while (Environment.TickCount - start < timeoutMs)
        {
            foreach (IntPtr tab in GetAllTabHandles(explorerTopLevel))
            {
                if (!before.Contains(tab))
                    return tab;
            }
            await Task.Delay(35, ct);
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
            dynamic dynWindows = windows;
            int count = 0;
            
            try
            {
                count = (int)dynWindows.Count;
            }
            catch (Exception ex) when (IsComOrRpcException(ex))
            {
                // The ShellWindows collection itself became invalid (e.g. Explorer restarted)
                return result;
            }

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
                catch (Exception ex) when (IsComOrRpcException(ex))
                {
                    // Usually 0x800706BA RPC Server Unavailable.
                    // The window died between enumeration and property access.
                    if (window is not null)
                        Marshal.FinalReleaseComObject(window);
                }
                catch
                {
                    if (window is not null)
                        Marshal.FinalReleaseComObject(window);
                }
            }
        }
        catch
        {
            // Catch anything else related to the dynamic binder or COM wrapper
        }
        finally
        {
            // Intentionally keep the cached ShellWindows collection alive.
        }

        return result;
    }

    private static bool IsComOrRpcException(Exception ex)
    {
        if (ex is COMException) return true;
        if (ex.InnerException is COMException) return true;
        if (ex.GetType().Name == "RuntimeBinderException") return true;
        return false;
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
                {
                    // Validate the cached object is still alive (Explorer may have restarted).
                    try
                    {
                        _ = ((dynamic)_shellWindows).Count;
                        return _shellWindows;
                    }
                    catch (Exception ex) when (IsComOrRpcException(ex))
                    {
                        // Stale COM object — discard and recreate below.
                        try
                        {
                            Marshal.FinalReleaseComObject(_shellWindows);
                        }
                        catch (Exception releaseEx)
                        {
                            Debug.WriteLine($"[ExplorerTabHookService] Failed to release stale ShellWindows COM object: {releaseEx.Message}");
                        }
                        _shellWindows = null;
                    }
                }

                // If called on a different thread, recreate to match COM apartment.
                if (_shellWindows is not null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(_shellWindows);
                    }
                    catch (Exception releaseEx)
                    {
                        Debug.WriteLine($"[ExplorerTabHookService] Failed to release ShellWindows COM object on thread switch: {releaseEx.Message}");
                    }
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

    private async Task<string?> WaitForRealLocation(object comTab, int timeoutMs, CancellationToken ct = default)
    {
        int start = Environment.TickCount;
        string? last = null;
        int stable = 0;

        while (Environment.TickCount - start < timeoutMs)
        {
            string? location = _shellComNavigator.TryGetComLocation(comTab);
            if (IsTabNavigableLocation(location))
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

            await Task.Delay(70, ct);
        }

        return null;
    }

    private static bool IsRealFileSystemLocation(string? location)
    {
        return OpenTargetClassifier.Classify(location).IsPhysicalFileSystem;
    }

    private static bool IsPhysicalFileSystemLocation(string? location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return false;

        if (location.StartsWith("\\\\", StringComparison.OrdinalIgnoreCase))
            return true;

        if (location.Length >= 3 && char.IsLetter(location[0]) && location[1] == ':' && (location[2] == '\\' || location[2] == '/'))
            return true;

        return false;
    }

    private static bool IsTabNavigableLocation(string? location)
    {
        if (IsRealFileSystemLocation(location))
            return true;

        return IsShellNamespaceLocation(location);
    }

    private static bool ShouldBypassAutoConvertForLocation(string? location)
    {
        return OpenTargetClassifier.Classify(location).RequiresNativeShellLaunch;
    }

    private static bool ShouldConvertWindowLocationToTab(string? location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return true;

        return !ShouldBypassAutoConvertForLocation(location);
    }

    private void RestoreEarlyHiddenWindow(IntPtr hwnd)
    {
        bool hadEarlyHide = _earlyHiddenExplorer.TryRemove(hwnd, out _);
        if (!hadEarlyHide)
            return;

        if (_windowManager.IsAlive(hwnd))
        {
            _windowManager.RestoreVisibility(hwnd);
            _windowManager.Show(hwnd);
        }
    }

    private static bool IsShellNamespaceLocation(string? location)
    {
        return OpenTargetClassifier.Classify(location).IsShellNamespace;
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

    private static string? TryGetNormalizedParentPath(string path)
    {
        try
        {
            string fullPath = Path.GetFullPath(path);
            string? parent = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrWhiteSpace(parent))
                return null;

            return NormalizePathForCompare(parent);
        }
        catch
        {
            return null;
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

    private void CleanupStaleDictionaryEntries()
    {
        foreach (IntPtr hwnd in _knownExplorerTopLevels.Keys)
        {
            if (!NativeMethods.IsWindow(hwnd))
                _knownExplorerTopLevels.TryRemove(hwnd, out _);
        }

        foreach (IntPtr hwnd in _earlyHiddenExplorer.Keys)
        {
            if (!NativeMethods.IsWindow(hwnd))
                _earlyHiddenExplorer.TryRemove(hwnd, out _);
        }

        foreach (IntPtr hwnd in _tabIndexCache.Keys)
        {
            if (!NativeMethods.IsWindow(hwnd))
                _tabIndexCache.TryRemove(hwnd, out _);
        }

        foreach (IntPtr hwnd in _newTabDefaultInFlight.Keys)
        {
            if (!NativeMethods.IsWindow(hwnd))
                _newTabDefaultInFlight.TryRemove(hwnd, out _);
        }

        foreach (IntPtr hwnd in _newTabDefaultSnapshots.Keys)
        {
            if (!NativeMethods.IsWindow(hwnd))
                _newTabDefaultSnapshots.TryRemove(hwnd, out _);
        }

        DateTimeOffset cutoff = DateTimeOffset.UtcNow.AddSeconds(-30);
        foreach (var kvp in _pending)
        {
            if (kvp.Value < cutoff)
                _pending.TryRemove(kvp.Key, out _);
        }

        foreach (var kvp in _suppressNewTabDefaultUntil)
        {
            if (kvp.Value <= DateTimeOffset.UtcNow)
                _suppressNewTabDefaultUntil.TryRemove(kvp.Key, out _);
        }

        foreach (var kvp in _externalSuppressNewTabDefaultUntil)
        {
            if (kvp.Value <= DateTimeOffset.UtcNow || !NativeMethods.IsWindow(kvp.Key))
                _externalSuppressNewTabDefaultUntil.TryRemove(kvp.Key, out _);
        }

        _nativeBrowseFallbackBypass.CleanupExpired();
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        _cts.Cancel();
        _cleanupTimer?.Stop();
        RestoreTrackedExplorerWindowsOnShutdown();

        _winEventHookManager.Dispose();

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

        _interceptOpenGate.Dispose();
        _cts.Dispose();
    }

    private void RestoreTrackedExplorerWindowsOnShutdown()
    {
        HashSet<IntPtr> trackedWindows =
        [
            .. _earlyHiddenExplorer.Keys,
            .. _pending.Keys
        ];

        foreach (IntPtr hwnd in trackedWindows)
        {
            try
            {
                if (!_windowManager.IsAlive(hwnd))
                    continue;

                _windowManager.RestoreVisibility(hwnd);
                if (!_windowManager.IsVisible(hwnd))
                    _windowManager.Show(hwnd);
            }
            catch (Exception ex)
            {
                _logger.Warn($"Failed to restore tracked Explorer window during shutdown: 0x{hwnd.ToInt64():X} ({ex.Message})");
            }
        }
    }

    public bool IsAutoConvertEnabled => _autoConvertEnabled;

    public void SetAutoConvertEnabled(bool enabled)
    {
        if (_disposed)
            return;

        bool previous = _autoConvertEnabled;
        if (previous == enabled)
            return;

        _autoConvertEnabled = enabled;
        _settings.EnableAutoConvertExplorerWindows = enabled;

        if (enabled)
            _logger.Info("Explorer auto-convert enabled at runtime.");
        else
            _logger.Info("Explorer auto-convert disabled at runtime.");
    }
}
