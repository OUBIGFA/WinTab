using Shell32;
using SHDocVw;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using WinTab.Helpers;
using WinTab.Interop;
using WinTab.Models;
using WinTab.WinAPI;

namespace WinTab.Hooks;

using WindowEntry = DualKeyEntry<InternetExplorer, nint?, WindowInfo>;

[Fody.ConfigureAwait(true)]
public partial class ExplorerWatcher : IHook
{
    private const int NavigationCompleteWaitMs = 600;
    private const int NavigationVerificationWaitMs = 1_200;
    /// <summary>
    /// How long a synchronous Explorer command may take to be acknowledged. A busy Explorer often needs
    /// more than a few hundred milliseconds to activate a tab; the outcome is always verified by
    /// observing the window afterwards, never by the acknowledgement alone.
    /// </summary>
    private const int ExplorerCommandTimeoutMs = 1_000;
    private const int StartupLocationCacheLimit = 512;
    private const int ClosedWindowHistoryLimit = 100;
    private static bool _instanceRunning;
    private static Guid _shellBrowserGuid = typeof(IShellBrowser).GUID;

    private ShellWindows _shellWindows = null!;
    private ShellPathComparer _shellPathComparer = null!;
    private readonly StaTaskScheduler _staTaskScheduler;
    private nint _mainWindowHandle;
    private readonly ConcurrentDictionary<nint, WindowIdentity> _processedHWnds = new();
    private readonly ConcurrentDictionary<nint, int> _hookedTopLevelUseCounts = new();
    private readonly ConcurrentDictionary<string, bool> _startupLocationCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly DualKeyDictionary<InternetExplorer, nint?, WindowInfo> _windowEntryDict = [];
    private readonly List<WindowRecord> _closedWindows = new();
    private readonly object _windowEntryDictLock = new(), _closedWindowsLock = new(), _processLock = new();
    private readonly SemaphoreSlim _toOpenWindowsLock = new(1);
    private readonly CoalescingAsyncWork _registrationWork;
    private readonly CoalescingAsyncWork _selectionWork;
    private readonly Func<int> _getDefaultExplorerLaunchId;
    private readonly Func<IEnumerable<nint>> _getExplorerWindows = ExplorerWindowDiscovery.GetAllExplorerWindows;
    private readonly ExplorerLaunchLocationResolver _locationResolver = new();
    private readonly ProcessWatcher _processWatcher;
    private int _mainExplorerProcessId;
    private Timer? _explorerCheckTimer;

    private WinEventHookThread? _winEventHookThread;
    private WinEventDelegate? _eventObjectShowHookCallback;
    private DShellWindowsEvents_WindowRegisteredEventHandler? _windowRegisteredHandler;

    private string _defaultLocation = null!;
    private volatile bool _reuseTabs = true;
    private volatile bool _isForcingTabs;
    private volatile bool _preExistingExplorerWindowsProtected;
    private readonly AsyncLocal<MergeOperation?> _currentMerge = new();
    private readonly MergeSourceConcealPulse _mergeSourceConcealPulse = new();
    private readonly ConcurrentDictionary<nint, ConcealedWindow> _mergeSourceHWnds = new();
    private readonly ConcurrentDictionary<nint, MergeOperation> _closingMergeSourceHWnds = new();

    internal TabStripHitTester TabStrip { get; } = new();
    public bool IsHookActive => _isForcingTabs;
    public bool IsShellReady => _mainExplorerProcessId != 0 && _shellWindows != null;
    public event Action? OnShellInitialized;

    public ExplorerWatcher(Func<int>? getDefaultExplorerLaunchId = null)
    {
        if (_instanceRunning)
            throw new InvalidOperationException("Only one instance of ExplorerWatcher is allowed at a time.");
        _instanceRunning = true;

        _staTaskScheduler = new StaTaskScheduler();
        _registrationWork = new CoalescingAsyncWork(() => RunShellWorkAsync(ProcessRegisteredShellWindowsAsync));
        _selectionWork = new CoalescingAsyncWork(() => RunShellWorkAsync(CacheActiveSelectionAsync));
        _mergeSafetyTimer = new Timer(RecoverExpiredMergeSources, null, Timeout.Infinite, Timeout.Infinite);
        _selectionTimer = new Timer(state => _selectionWork.Request(), null, Timeout.Infinite, Timeout.Infinite);
        _getDefaultExplorerLaunchId = getDefaultExplorerLaunchId ?? (static () => 1);
        _processWatcher = new ProcessWatcher("explorer");
        _processWatcher.ProcessTerminated += OnExplorerProcessTerminated;
        StartExplorerProcessCheck();
    }

    public void StartHook()
    {
        if (_isForcingTabs || _disposed) return;
        RecoverHiddenExplorerWindows("start-hook");
        Interlocked.Increment(ref _hookGeneration);
        _hookLifetime.Dispose();
        _hookLifetime = new CancellationTokenSource();
        _isForcingTabs = true;
        ScheduleShellWindowRegistration();
        ConcealPreloadedExplorerFrames();
        ExplorerDebugLog.Write("StartHook");
    }

    public void StopHook()
    {
        _isForcingTabs = false;
        Interlocked.Increment(ref _hookGeneration);
        _hookLifetime.Cancel();
        StopMergeSourceConcealPulse();
        RecoverHiddenExplorerWindows("stop-hook");
    }
    public void SetReuseTabs(bool reuseTabs) => _reuseTabs = reuseTabs;

    private bool TrySearchForTab(string targetPath, nint excludedTopLevelWindow, out nint tabHandle,
        out InternetExplorer? foundWindow)
    {
        nint targetPidl = 0;
        tabHandle = 0;
        foundWindow = null;
        try
        {
            var normalizedTargetPath = Helper.NormalizeLocation(targetPath);
            var targetPidlAttempted = false;
            var candidates = new List<ExplorerTabReuseCandidate>();
            var candidateOwners = new Dictionary<nint, (InternetExplorer Window, WindowInfo Info, WindowIdentity TabIdentity)>();

            // Hold the lock for the whole scan: concurrent .Add/.Remove during enumeration would
            // throw and the outer catch would silently fail the search.
            lock (_windowEntryDictLock)
            {
                foreach (var (window, windowInfo, _) in ((IEnumerable<WindowEntry>)_windowEntryDict).ToArray())
                {
                    if (!TryGetKnownTabHandle(window, out var tab) || windowInfo.Closed)
                        continue;

                    var topLevelWindow = windowInfo.Identity.Handle;
                    if (excludedTopLevelWindow != 0 && topLevelWindow == excludedTopLevelWindow)
                        continue;
                    if (!candidateOwners.TryAdd(tab, (window, windowInfo, windowInfo.TabIdentity)))
                        continue;

                    candidates.Add(new ExplorerTabReuseCandidate(
                        tab,
                        windowInfo.Location,
                        () => TryGetLocation(window),
                        location => windowInfo.Location = location));
                }
            }

            bool AreEquivalent(string left, string right)
            {
                if (StringComparer.OrdinalIgnoreCase.Equals(left, right))
                    return true;

                if (StringComparer.OrdinalIgnoreCase.Equals(
                        normalizedTargetPath,
                        Helper.NormalizeLocation(right)))
                    return true;

                if (!targetPidlAttempted)
                {
                    targetPidlAttempted = true;
                    targetPidl = _shellPathComparer.GetPidlFromPath(left);
                }

                return targetPidl != 0 && _shellPathComparer.IsEquivalent(left, right, targetPidl);
            }

            if (!ExplorerTabReuseMatcher.TryFind(targetPath, candidates, AreEquivalent, out var matchedTabHandle))
                return false;

            lock (_windowEntryDictLock)
            {
                if (!_windowEntryDict.TryGetValue(matchedTabHandle, out foundWindow) || foundWindow == null)
                    return false;
                if (!_windowEntryDict.TryGetValue(foundWindow, out WindowInfo? info) ||
                    !IsCurrentWindow(foundWindow, info) || !IsCurrentTab(info, matchedTabHandle))
                    return false;
                var owner = candidateOwners[matchedTabHandle];
                if (!ReferenceEquals(foundWindow, owner.Window) || !ReferenceEquals(info, owner.Info) || !owner.TabIdentity.IsCurrent)
                    return false;
            }

            tabHandle = matchedTabHandle;
            return true;
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            ExplorerDebugLog.Write($"Tab search temporarily unavailable error={exception.HResult:X8}");
            tabHandle = 0;
            return false;
        }
        finally
        {
            if (targetPidl != 0)
                Marshal.FreeCoTaskMem(targetPidl);
        }
    }
    public async Task<bool> SelectTabByHandle(nint windowHandle, nint tabHandle, int timeoutMs = 2_500)
    {
        if (windowHandle == 0 || tabHandle == 0)
            return false;
        try
        {
            var parentIdentity = WindowIdentity.Capture(windowHandle);
            var tabIdentity = WindowIdentity.Capture(tabHandle);
            EnsureWindowIdentity(parentIdentity);
            EnsureWindowIdentity(tabIdentity);
            Helper.RestoreWindowToForeground(windowHandle);
            var selected = await TabSelectionEngine.CycleToTabAsync(tabHandle,
                () => ExplorerWindowDiscovery.GetAllExplorerTabs(windowHandle).ToArray(),
                () => GetActiveTabHandle(windowHandle),
                index =>
                {
                    EnsureWindowIdentity(parentIdentity);
                    EnsureWindowIdentity(tabIdentity);
                    SelectTabByIndex(windowHandle, index);
                },
                totalTimeoutMs: timeoutMs, perStepTimeoutMs: 250, cancellationToken: CurrentCancellation);
            EnsureWindowIdentity(parentIdentity);
            EnsureWindowIdentity(tabIdentity);
            return selected;
        }
        catch (TimeoutException)
        {
            ExplorerDebugLog.Write($"Tab switch timed out hwnd={windowHandle}");
            return false;
        }
        catch (OperationCanceledException)
        {
            return false;
        }
    }

    internal void SelectLastTab(nint windowHandle)
    {
        var count = ExplorerWindowDiscovery.GetAllExplorerTabs(windowHandle).Count();
        if (count > 0)
            WinApi.TrySendMessage(windowHandle, WinApi.WM_COMMAND, 0xA221, count, ExplorerCommandTimeoutMs);
    }

    private void SelectTabByIndex(nint windowHandle, int index)
    {
        EnsureCurrentMerge();
        // A slow acknowledgement is not a failed switch: the caller keeps observing the active tab.
        WinApi.TrySendMessage(windowHandle, WinApi.WM_COMMAND, 0xA221, index + 1, ExplorerCommandTimeoutMs);
    }
    private async Task RequestToOpenNewTab(nint windowHandle, bool bringToFront = false, bool lockToOpenWindows = true)
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
        EnsureCurrentMerge();
        if (!WinApi.PostMessage(tabHandle, WinApi.WM_COMMAND, 0xA21B, 0))
            throw new InvalidOperationException("Explorer did not accept the new-tab command.");

        if (bringToFront)
            Helper.RestoreWindowToForeground(windowHandle);
    }

    private bool IsWindowProtected(nint handle) =>
        _processedHWnds.TryGetValue(handle, out var identity) && identity.IsCurrent;

    private void PreventWindowHiding(nint handle)
    {
        var identity = WindowIdentity.Capture(handle);
        if (!identity.IsCurrent)
            return;
        if (_processedHWnds.TryGetValue(handle, out var current) && current == identity)
            return;
        _processedHWnds[handle] = identity;
        _ = Task.Delay(7_000).ContinueWith(completed =>
            _processedHWnds.TryRemove(new KeyValuePair<nint, WindowIdentity>(handle, identity)), TaskScheduler.Default);
    }
    private void OnWindowShown(nint hWinEventHook, uint eventType, nint hWnd, int idObject, int idChild, uint dwEventThread, uint dWmsEventTime)
    {
        if (!_isForcingTabs || !_preExistingExplorerWindowsProtected || _disposed || hWnd == 0) return;

        // OBJID_WINDOW = 0 and CHILDID_SELF = 0. The system-wide WinEvent hook range fires for every
        // accessibility sub-element on the desktop (caret, focus, menu items, scrollbars, list items,
        // alerts, etc.). Without this filter every caret blink in every program would push us into the
        // expensive Explorer-top-level lookup + ShellWindows registration scheduler + 25ms-pulse worker.
        if (idObject != 0 || idChild != 0) return;

        // Skip non-Explorer events entirely so the conceal pulse and registration scheduler only fire
        // when something actually touched a CabinetWClass window.
        var explorerTopLevel = GetExplorerTopLevelWindow(hWnd);
        if (explorerTopLevel == 0) return;

        TryHideIncomingExplorerWindow(explorerTopLevel);
        StartMergeSourceConcealPulse();
        ScheduleShellWindowRegistration(1);
    }
    private bool TryHideIncomingExplorerWindow(nint hWnd)
    {
        hWnd = GetExplorerTopLevelWindow(hWnd);
        if (!_isForcingTabs || hWnd == 0) return false;
        if (_closingMergeSourceHWnds.ContainsKey(hWnd))
        {
            HideMergeSourceWindow(hWnd);
            RequestCloseMergedSourceWindow(hWnd);
            return true;
        }

        if (IsWindowProtected(hWnd)) return false;
        if (Helper.IsCtrlShiftDown()) return false;
        if (_hookedTopLevelUseCounts.ContainsKey(hWnd)) return false;
        if (_mainWindowHandle != 0 && hWnd == _mainWindowHandle) return false;
        if (ExplorerWindowDiscovery.GetAllExplorerTabs(hWnd).Take(2).Count() > 1) return false;

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
        if (!_isForcingTabs || hWnd == 0) return;
        if (IsWindowProtected(hWnd)) return;
        if (Helper.IsCtrlShiftDown()) return;
        if (_mainWindowHandle != 0 && hWnd == _mainWindowHandle) return;
        if (ExplorerWindowDiscovery.GetAllExplorerTabs(hWnd).Take(2).Count() > 1) return;

        var targetWindow = GetMainWindowHWnd(hWnd);
        if (targetWindow == 0 || hWnd == targetWindow) return;

        HideMergeSourceWindow(hWnd);
    }
    /// <summary>
    /// Windows 11 preloads a hidden Explorer frame and reuses it for the next folder the user opens.
    /// Conceal such frames up front so that reuse never flashes on screen; their merge budget only
    /// starts once Explorer actually shows them.
    /// </summary>
    private void ConcealPreloadedExplorerFrames()
    {
        if (!_isForcingTabs || !_preExistingExplorerWindowsProtected || _disposed)
            return;

        foreach (var handle in _getExplorerWindows())
        {
            if (!WinApi.IsWindowVisible(handle))
                TryHideIncomingExplorerWindow(handle);
        }
    }
    private void HideMergeSourceWindow(nint handle)
    {
        if (!_isForcingTabs || !_preExistingExplorerWindowsProtected || _disposed || handle == 0 ||
            (_currentMerge.Value is { } operation && !operation.IsCurrent))
            return;

        if (_mergeSourceHWnds.TryGetValue(handle, out var existing) && !existing.Identity.IsCurrent)
            _mergeSourceHWnds.TryRemove(new KeyValuePair<nint, ConcealedWindow>(handle, existing));
        var concealed = _mergeSourceHWnds.GetOrAdd(handle,
            windowHandle => new ConcealedWindow(WindowIdentity.Capture(windowHandle), _hookGeneration));
        StartMergeBudgetIfShown(concealed);
        ConcealMergeSourceWindow(concealed);
        _mergeSafetyTimer.Change(250, 250);
    }

    private async Task RestoreMergeSourceWindowAsync(nint handle)
    {
        if (!_mergeSourceHWnds.TryGetValue(handle, out var concealed))
            return;
        if (_currentMerge.Value is { } operation &&
            (operation.Identity != concealed.Identity || operation.Generation != concealed.Generation))
            return;
        var restored = await Helper.DoUntilConditionAsync(() => RestoreConcealedWindow(concealed),
            result => result, 1_000, 50);
        if (!restored)
            ReportStatus("Explorer rejected window recovery; recovery will be retried.");
    }

    private void RemoveMergeSourceTracking(nint handle)
    {
        if (!_mergeSourceHWnds.TryGetValue(handle, out var concealed))
            return;
        if (_currentMerge.Value is { } operation && operation.Identity != concealed.Identity)
            return;
        _mergeSourceHWnds.TryRemove(new KeyValuePair<nint, ConcealedWindow>(handle, concealed));
        RemoveClosingMergeSource(concealed.Identity);
        ExplorerWindowVisibility.Forget(concealed.Identity);
        if (concealed.Identity.IsCurrent)
            PreventWindowHiding(handle);
    }

    private void StartMergeSourceConcealPulse(int durationMs = 1_200)
    {
        _mergeSourceConcealPulse.Start(
            () => _isForcingTabs && _preExistingExplorerWindowsProtected && !_mergeSourceHWnds.IsEmpty,
            ConcealMergeSourceWindowsOnce, durationMs);
    }

    private void ConcealMergeSourceWindowsOnce()
    {
        foreach (var concealed in _mergeSourceHWnds.Values)
            ConcealMergeSourceWindow(concealed);
    }

    private void ConcealMergeSourceWindow(ConcealedWindow concealed)
    {
        lock (concealed)
        {
            if (!_disposed && !concealed.Recovering && _isForcingTabs &&
                concealed.Generation == _hookGeneration && concealed.Identity.IsCurrent &&
                Environment.TickCount64 < concealed.ExpiresAt)
                ExplorerWindowVisibility.Hide(concealed.Identity);
        }
    }

    private void StopMergeSourceConcealPulse() => _mergeSourceConcealPulse.Stop();

    /// <summary>
    /// Finds the first tracked ShellWindow whose top-level Explorer window is <paramref name="hWnd"/>
    /// and satisfies <paramref name="predicate"/>. COM reads are guarded because tracked windows can
    /// be torn down by Explorer at any time.
    /// </summary>
    private InternetExplorer? FindTrackedWindowByTopLevel(nint hWnd, Func<InternetExplorer, WindowInfo, bool>? predicate = null)
    {
        lock (_windowEntryDictLock)
        {
            foreach (var (window, info, _) in _windowEntryDict)
            {
                try
                {
                    if (info.Identity.Handle != hWnd || !info.Identity.IsCurrent)
                        continue;

                    if (predicate == null || predicate(window, info))
                        return window;
                }
                catch
                {
                    // 忽略无法读取的窗口状态。
                }
            }
        }

        return null;
    }
    private bool HasTrackedTopLevelWindow(nint hWnd) => FindTrackedWindowByTopLevel(hWnd) != null;
    private bool HasHookedShellWindowForTopLevel(nint hWnd) => FindTrackedWindowByTopLevel(hWnd, (_, info) => info.EventsHooked) != null;
    private bool HasOtherTrackedShellWindowForTopLevel(InternetExplorer currentWindow, nint hWnd) =>
        FindTrackedWindowByTopLevel(hWnd, (window, _) => !ReferenceEquals(window, currentWindow)) != null;
    private bool HasNonStartupShellWindowForTopLevel(nint hWnd) =>
        FindTrackedWindowByTopLevel(hWnd, (window, info) => !IsStartupExplorerLocation(info.Location ?? string.Empty)) != null;

    private void ReleaseTopLevelTrackingIfUnused(nint hWnd)
    {
        if (hWnd == 0)
            return;

        if (HasTrackedTopLevelWindow(hWnd) || ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd))
        {
            _ = Task.Delay(5_000).ContinueWith(_ =>
            {
                if (!HasTrackedTopLevelWindow(hWnd) && !ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd))
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
    private List<(InternetExplorer Window, WindowInfo WindowInfo, bool IsNewTopLevel)> AdoptNewShellWindows()
    {
        RemoveExpiredWindowEntries();
        var result = new List<(InternetExplorer Window, WindowInfo WindowInfo, bool IsNewTopLevel)>();
        var singleTabTopLevelsInBatch = new HashSet<nint>();
        var count = _shellWindows.Count;

        for (var index = count - 1; index >= 0; index--)
        {
            try
            {
                if (_shellWindows.Item(index) is not InternetExplorer window)
                    continue;

                WindowInfo windowInfo;
                nint hWnd;
                bool wasTrackedTopLevel;
                lock (_windowEntryDictLock)
                {
                    if (_windowEntryDict.TryGetValue(window, out WindowEntry tracked))
                    {
                        if (!tracked.Value.EventsHooked || !tracked.OptionalKey.HasValue ||
                            !IsCurrentTab(tracked.Value, tracked.OptionalKey.Value))
                            result.Add((window, tracked.Value, false));
                        continue;
                    }

                    windowInfo = CreateWindowInfo(window);
                    if (!windowInfo.Identity.IsCurrent)
                    {
                        _registrationRetryPending = true;
                        continue;
                    }
                    hWnd = windowInfo.Identity.Handle;
                    wasTrackedTopLevel = HasTrackedTopLevelWindow(hWnd);
                    var tabCount = ExplorerWindowDiscovery.GetAllExplorerTabs(hWnd).Take(2).Count();
                    if (tabCount <= 1 &&
                        (wasTrackedTopLevel || !singleTabTopLevelsInBatch.Add(hWnd)))
                        continue;

                    if (!wasTrackedTopLevel &&
                        !IsWindowProtected(hWnd) &&
                        _isForcingTabs &&
                        !Helper.IsCtrlShiftDown() &&
                        _mainWindowHandle != hWnd)
                    {
                        HideMergeSourceWindow(hWnd);
                    }

                    _windowEntryDict.Add(window, windowInfo);
                    if (_windowEntryDict.Count == 1)
                        _mainWindowHandle = hWnd;
                }

                if (!wasTrackedTopLevel)
                    TryHideRegisteredMergeSourceWindow(hWnd);
                result.Add((window, windowInfo, !wasTrackedTopLevel));
            }
            catch (COMException exception) when (!IsDisconnectedShell(exception))
            {
                _registrationRetryPending = true;
                ExplorerDebugLog.Write($"Shell window registration deferred error={exception.GetType().Name}");
            }
        }

        return result;
    }
    private void OnShellWindowRegistered(int cookie)
    {
        ScheduleShellWindowRegistration();
    }

    private void ScheduleShellWindowRegistration(int delayMs = 25)
    {
        if (!_disposed && _preExistingExplorerWindowsProtected)
            _registrationWork.Request();
    }

    private async Task ProcessRegisteredShellWindowsAsync()
    {
        if (_shellWindows == null || _disposed)
            return;

        var generation = _shellGeneration;
        try
        {
            _registrationRetryPending = false;
            await Task.Delay(1, _shellLifetime.Token);
            for (var attempt = 0; attempt < 4; attempt++)
            {
                if (generation != _shellGeneration || _disposed)
                    return;
                var windows = AdoptNewShellWindows();
                if (windows.Count == 0)
                {
                    if (attempt == 0)
                    {
                        await Task.Delay(50, _shellLifetime.Token);
                        continue;
                    }
                    return;
                }
                await Task.WhenAll(windows.Select(async item =>
                {
                    if (item.IsNewTopLevel)
                    {
                        await ProcessRegisteredShellWindowAsync(item.Window, item.WindowInfo);
                        return;
                    }
                    await RestoreMergeSourceWindowAsync(item.WindowInfo.Identity.Handle);
                    await RegisterIndependentWindowAsync(item.Window, item.WindowInfo, item.WindowInfo.Identity.Handle);
                }));
            }
        }
        catch (OperationCanceledException)
        {
            ExplorerDebugLog.Write("Registration cancelled");
        }
        catch (Exception exception) when (IsDisconnectedShell(exception))
        {
            RetireShellConnection("shell-connection-lost");
            ReportStatus($"Explorer catalog disconnected ({exception.HResult:X8}); reconnecting without restarting Explorer.");
        }
        catch (Exception exception)
        {
            _registrationRetryPending = true;
            ExplorerDebugLog.Write($"Registration failed error={exception.GetType().Name}");
            RecoverHiddenExplorerWindows("registration-failed");
            ReportStatus("Explorer registration failed; source windows were restored.");
        }
    }
    private async Task ProcessRegisteredShellWindowAsync(InternetExplorer window, WindowInfo windowInfo)
    {
        var showAgain = true;
        var removed = false;
        nint hWnd = windowInfo.Identity.Handle;
        var previousOperation = _currentMerge.Value;
        var hookGeneration = _hookGeneration;
        using var operation = new MergeOperation(windowInfo.Identity, hookGeneration, _shellLifetime.Token,
            () => _isForcingTabs && hookGeneration == _hookGeneration && IsCurrentWindow(window, windowInfo),
            RemainingMergeTime(hWnd), _hookLifetime.Token);
        _currentMerge.Value = operation;

        try
        {
            if (!IsCurrentWindow(window, windowInfo))
                return;
            // Windows aggressively reuses hwnds. A hwnd that was recently merged-and-closed
            // gets a PreventWindowHiding 7s grace; if explorer.exe assigns the same hwnd to a
            // brand-new window during that window, we must still adopt the new COM object as
            // an independent tracked window. Returning early here used to leave the new window
            // in _windowEntryDict unhooked — once the 7s expired, the next WinEvent would hide
            // it as a merge source and nothing would ever restore it (the transparent 此电脑
            // residual).
            if (IsWindowProtected(hWnd))
            {
                ExplorerDebugLog.Write($"Registered already-processed hwnd={hWnd}");
                await RestoreMergeSourceWindowAsync(hWnd);
                await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                return;
            }

            if (Helper.IsCtrlShiftDown())
            {
                ExplorerDebugLog.Write($"Registered release ctrl-shift hwnd={hWnd}");
                await RestoreMergeSourceWindowAsync(hWnd);
                await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                return;
            }

            if (HasOtherTrackedShellWindowForTopLevel(window, hWnd))
            {
                ExplorerDebugLog.Write($"Registered sibling hwnd={hWnd}");
                await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                return;
            }

            var targetWindow = GetMainWindowHWnd(hWnd);
            if (!_isForcingTabs)
            {
                ExplorerDebugLog.Write($"Registered release disabled hwnd={hWnd}");
                await RestoreMergeSourceWindowAsync(hWnd);
                await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                return;
            }

            if (IsWindowProtected(hWnd))
            {
                ExplorerDebugLog.Write($"Registered already-processed late hwnd={hWnd}");
                await RestoreMergeSourceWindowAsync(hWnd);
                await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                return;
            }

            EnsureCurrentMerge();
            HideMergeSourceWindow(hWnd);

            var location = await ResolveInitialLocation(window);
            EnsureCurrentMerge();
            if (!string.IsNullOrWhiteSpace(location))
                windowInfo.Location = location;
            ExplorerDebugLog.Write($"Registered resolved hwnd={hWnd} target={targetWindow} location={location}");
            if (string.IsNullOrWhiteSpace(location) ||
                location.StartsWith("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", StringComparison.OrdinalIgnoreCase))
            {
                ExplorerDebugLog.Write($"Registered release unsupported-location hwnd={hWnd} location={location}");
                await RestoreMergeSourceWindowAsync(hWnd);
                await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                return;
            }

            var sourceAlive = windowInfo.Identity.IsCurrent;
            if (sourceAlive && !IsStartupExplorerLocation(location))
            {
                _ = await GetTabHandle(window);
                var tabCount = await WaitForExplorerTabCount(hWnd);
                EnsureCurrentMerge();
                if (tabCount != 1)
                {
                    ExplorerDebugLog.Write($"Registered release tab-count hwnd={hWnd} count={tabCount}");
                    await RestoreMergeSourceWindowAsync(hWnd);
                    await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                    return;
                }
            }

            targetWindow = GetMainWindowHWnd(hWnd, location);
            if (targetWindow == 0 || hWnd == targetWindow)
            {
                ExplorerDebugLog.Write($"Registered release no-target hwnd={hWnd} location={location}");
                await RestoreMergeSourceWindowAsync(hWnd);
                await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                return;
            }

            WindowRecord? recentlyClosedWindow = null;
            if (sourceAlive && TryGetRecentlyClosedWindow(location, out var closedWindow))
            {
                recentlyClosedWindow = closedWindow;
                ExplorerDebugLog.Write($"Registered merge recently-closed hwnd={hWnd} location={location}");
            }

            if (sourceAlive && !IsStartupExplorerLocation(location))
                HideMergeSourceWindow(hWnd);

            windowInfo.RefreshSelection(() => TryGetSelectedItems(window), () => IsCurrentWindow(window, windowInfo));
            EnsureCurrentMerge();
            var selectedItems = windowInfo.SelectedItems;
            if ((selectedItems == null || selectedItems.Length == 0) &&
                recentlyClosedWindow?.SelectedItems?.Length > 0)
            {
                selectedItems = recentlyClosedWindow.SelectedItems;
            }

            var record = new WindowRecord(location, hWnd, selectedItems);
            if (!await OpenTabNavigateWithSelection(record, targetWindow))
            {
                ExplorerDebugLog.Write($"Registered merge-failed hwnd={hWnd} location={location}");
                if (ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd))
                {
                    await RestoreMergeSourceWindowAsync(hWnd);
                    await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                }
                else
                {
                    ExplorerDebugLog.Write($"Registered merge-failed source already closed; not reopening intermediate hwnd={hWnd} location={location}");
                    RemoveMergeSourceTracking(hWnd);
                    RemoveWindowAndUnhookEvents(window, windowInfo, restoreHiddenWindow: false);
                    removed = true;
                }
                return;
            }

            ExplorerDebugLog.Write($"Registered merge-succeeded hwnd={hWnd} location={location}");
            UnhookWindowEvents(window, windowInfo);
            if (await CloseMergedSourceWindowAsync(window, hWnd))
            {
                showAgain = false;
                RemoveMergeSourceTracking(hWnd);
                RemoveWindowAndUnhookEvents(window, windowInfo, restoreHiddenWindow: false);
                removed = true;
            }
            else
            {
                await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
            }
        }
        catch (OperationCanceledException)
        {
            ExplorerDebugLog.Write($"Merge cancelled or timed out hwnd={hWnd}");
        }
        catch (Exception ex) when (!IsDisconnectedShell(ex))
        {
            ExplorerDebugLog.Write($"Registered error hwnd={hWnd} error={ex.GetType().Name}:{ex.Message}");
        }
        finally
        {
            try
            {
                if (!removed && showAgain && hWnd != 0)
                    await RestoreMergeSourceWindowAsync(hWnd);
                _currentMerge.Value = previousOperation;
                if (!removed && IsCurrentWindow(window, windowInfo))
                    await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
            }
            catch (Exception exception) when (!IsDisconnectedShell(exception))
            {
                ExplorerDebugLog.Write($"Window release failed error={exception.GetType().Name}");
            }
            finally
            {
                _currentMerge.Value = previousOperation;
            }
        }
    }
    private async Task RegisterIndependentWindowAsync(InternetExplorer window, WindowInfo windowInfo, nint hWnd)
    {
        if (!IsCurrentWindow(window, windowInfo))
            return;
        if (hWnd != 0)
            PreventWindowHiding(hWnd);

        HookWindowEvents(window, windowInfo);
        if (!IsCurrentWindow(window, windowInfo))
            return;

        var tabHandle = await ExplorerTabHandleResolver.WaitAsync(
            () => GetTabHandle(window), timeoutMs: 2_000, pollSleepMs: 50);
        if (tabHandle == 0 && IsCurrentWindow(window, windowInfo))
            _registrationRetryPending = true;
    }
    private async Task<bool> CloseMergedSourceWindowAsync(InternetExplorer window, nint handle)
    {
        var operation = _currentMerge.Value;
        if (operation == null)
            return false;
        var identity = operation.Identity;
        if (!identity.IsCurrent)
            return true;
        using var completion = CreateMergeCompletionOperation(window, operation);
        if (completion == null)
            return false;
        _currentMerge.Value = completion;
        _closingMergeSourceHWnds[handle] = completion;
        try
        {
            for (var attempt = 0; attempt < 2; attempt++)
            {
                EnsureCurrentMerge();
                if (!RequestCloseMergedSourceWindow(handle))
                    return !identity.IsCurrent;
                var closed = await Helper.DoUntilConditionAsync(() => !identity.IsCurrent,
                    isClosed => isClosed, attempt == 0 ? 700 : 300, 40, CurrentCancellation);
                if (closed)
                    return true;
            }
            ReportStatus("Explorer did not close the source window; it has been restored.");
            return false;
        }
        finally
        {
            _closingMergeSourceHWnds.TryRemove(new KeyValuePair<nint, MergeOperation>(handle, completion));
            try
            {
                if (identity.IsCurrent)
                    await RestoreMergeSourceWindowAsync(handle);
            }
            finally
            {
                _currentMerge.Value = operation;
            }
        }
    }

    private bool RequestCloseMergedSourceWindow(nint handle)
    {
        if (!_isForcingTabs || !_closingMergeSourceHWnds.TryGetValue(handle, out var operation) || !operation.IsCurrent ||
            ExplorerWindowDiscovery.GetAllExplorerTabs(handle).Take(2).Count() > 1)
            return false;
        // Explorer destroys the frame while handling WM_CLOSE, which takes longer than any acknowledgement
        // wait and makes a synchronous send report a failure for a successful close. Post the request and
        // let the caller observe whether the window went away.
        var posted = WinApi.PostMessage(handle, WinApi.WM_CLOSE, 0, 0);
        if (!posted)
            ExplorerDebugLog.Write($"Source close command failed hwnd={handle}");
        return posted;
    }
    private Task<string> ResolveInitialLocation(InternetExplorer window)
    {
        return _locationResolver.ResolveAsync(
            () => TryGetLocation(window),
            IsStartupExplorerLocation,
            isBusy: () => IsShellWindowBusy(window), cancellationToken: CurrentCancellation);
    }
    private static bool IsShellWindowBusy(InternetExplorer window)
    {
        try
        {
            return window.Busy;
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            return false;
        }
    }
    private static string TryGetLocation(InternetExplorer window)
    {
        try
        {
            return GetLocation(window);
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            ExplorerDebugLog.Write($"Tab location temporarily unavailable error={exception.HResult:X8}");
            return string.Empty;
        }
    }
    private static nint SafeGetWindowHandle(InternetExplorer window)
    {
        try
        {
            return new IntPtr(window.HWND);
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
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
    private Task<int> WaitForExplorerTabCount(nint hWnd)
    {
        return Helper.DoUntilConditionAsync(
            () => ExplorerWindowDiscovery.GetAllExplorerTabs(hWnd).Take(2).Count(),
            count => count > 0,
            800,
            20, CurrentCancellation);
    }
    private void HookWindowEvents(InternetExplorer window, WindowInfo windowInfo)
    {
        if (windowInfo.EventsHooked || !IsCurrentWindow(window, windowInfo))
            return;

        var hookedTopLevelHWnd = SafeGetWindowHandle(window);
        if (hookedTopLevelHWnd != 0)
        {
            windowInfo.HookedTopLevelHWnd = hookedTopLevelHWnd;
            _hookedTopLevelUseCounts.AddOrUpdate(hookedTopLevelHWnd, 1, (_, count) => count + 1);
        }

        // Create a strongly-typed handler so we can remove it later
        windowInfo.OnQuitHandler = () =>
        {
            try
            {
                windowInfo.RefreshSelection(() => TryGetSelectedItems(window), () => IsCurrentWindow(window, windowInfo));
            }
            catch (Exception exception) when (IsDisconnectedShell(exception))
            {
                ExplorerDebugLog.Write($"Closing tab selection already disconnected error={exception.HResult:X8}");
            }
            // Remember real folders so a quick re-open of the same folder can restore its selection.
            // Home, This PC, etc. carry nothing worth restoring.
            var location = windowInfo.Location;
            if (!string.IsNullOrWhiteSpace(location) && location != _defaultLocation)
                RememberClosedWindow(new WindowRecord(location, windowInfo.Identity.Handle, windowInfo.SelectedItems));

            RemoveWindowAndUnhookEvents(window, windowInfo);
        };
        windowInfo.OnNavigateHandler = (object _, ref object url) =>
        {
            if (!IsCurrentWindow(window, windowInfo))
                return;
            var location = url?.ToString();
            if (!string.IsNullOrWhiteSpace(location))
                windowInfo.Location = Helper.NormalizeLocation(location);
            windowInfo.SelectedItems = null;
            _selectionWork.Request();
        };

        try
        {
            if (string.IsNullOrWhiteSpace(windowInfo.Location))
                windowInfo.Location = TryGetLocation(window);

            window.OnQuit += windowInfo.OnQuitHandler;
            window.NavigateComplete2 += windowInfo.OnNavigateHandler;
            windowInfo.EventsHooked = true;

            // Make sure the window is still alive (User might have closed it immediately after opening it)
            var hWnd = new IntPtr(window.HWND);
            if (ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd))
                TabStrip.ScheduleRefresh(hWnd);
        }
        catch (Exception exception)
        {
            RemoveWindowAndUnhookEvents(window, windowInfo);
            if (IsDisconnectedShell(exception))
                throw;
            ExplorerDebugLog.Write($"Window event registration failed error={exception.GetType().Name}:{exception.Message}");
        }
    }
    private void UnhookWindowEvents(InternetExplorer window, WindowInfo windowInfo)
    {
        var onQuit = windowInfo.OnQuitHandler;
        var onNavigate = windowInfo.OnNavigateHandler;
        windowInfo.OnQuitHandler = null;
        windowInfo.OnNavigateHandler = null;
        windowInfo.EventsHooked = false;
        ReleaseHookedTopLevel(windowInfo);

        if (onQuit != null)
            Detach(() => window.OnQuit -= onQuit);
        if (onNavigate != null)
            Detach(() => window.NavigateComplete2 -= onNavigate);

        static void Detach(Action detach)
        {
            try
            {
                detach();
            }
            catch (Exception exception) when (exception is COMException or InvalidComObjectException ||
                exception is System.Reflection.TargetInvocationException { InnerException: COMException or InvalidComObjectException })
            {
                ExplorerDebugLog.Write($"Window event connection already closed error={exception.GetType().Name}");
            }
        }
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
        if (!IsRegisteredWindow(window, windowInfo))
            return;
        windowInfo.Closed = true;

        // Remove from dictionary
        if (useLock)
        {
            lock (_windowEntryDictLock)
                _windowEntryDict.Remove(window);
        }
        else
            _windowEntryDict.Remove(window);

        UnhookWindowEvents(window, windowInfo);

        try
        {
            var hWnd = windowInfo.Identity.Handle;
            if (_closingMergeSourceHWnds.ContainsKey(hWnd))
                restoreHiddenWindow = false;

            if (restoreHiddenWindow)
                ExplorerWindowVisibility.Restore(windowInfo.Identity);

            if (_closingMergeSourceHWnds.ContainsKey(hWnd))
                RemoveMergeSourceTracking(hWnd);

            _processedHWnds.TryRemove(new KeyValuePair<nint, WindowIdentity>(hWnd, windowInfo.Identity));
            ReleaseTopLevelTrackingIfUnused(hWnd);
            TabStrip.Forget(hWnd);
            if (_mainWindowHandle == hWnd && !ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd))
                _mainWindowHandle = 0;
        }
        catch
        {
            // 忽略主窗口句柄重置失败，继续释放 COM 引用。
        }

        // Finally, release the COM reference for this InternetExplorer instance
        if (Marshal.IsComObject(window))
        {
            try
            {
                Marshal.ReleaseComObject(window);
            }
            catch (InvalidComObjectException exception)
            {
                ExplorerDebugLog.Write($"Window connection already released error={exception.GetType().Name}");
            }
        }
    }
    private void RememberClosedWindow(WindowRecord record)
    {
        lock (_closedWindowsLock)
        {
            _closedWindows.Add(record);
            if (_closedWindows.Count > ClosedWindowHistoryLimit)
                _closedWindows.RemoveRange(0, _closedWindows.Count - ClosedWindowHistoryLimit);
        }
    }

    private async Task OpenNewWindowWithSelection(WindowRecord windowToOpen, bool lockToOpenWindows = true)
    {
        if (lockToOpenWindows)
            await _toOpenWindowsLock.WaitAsync(CurrentCancellation);

        try
        {
            // The new window registers through ShellWindows a moment later; keeping the record lets
            // that registration pick up the requested selection.
            RememberClosedWindow(windowToOpen);

            var hasSelection = windowToOpen.SelectedItems?.Length > 0;

            nint[]? currentWindows = null;
            if (hasSelection)
                currentWindows = ExplorerWindowDiscovery.GetAllExplorerWindows().ToArray();

            Helper.BypassWinForegroundRestrictions();

            var location = string.IsNullOrWhiteSpace(windowToOpen.Location) ? _defaultLocation : windowToOpen.Location;
            await RunInStaThread(() =>
            {
                Shell? shell = null;
                try
                {
                    shell = new Shell();
                    EnsureCurrentMerge();
                    shell.ShellExecute(location, "", "", "opennewwindow");
                }
                finally
                {
                    if (shell != null)
                        Marshal.ReleaseComObject(shell);
                }
            });

            if (!hasSelection) return;

            var newWindowHandle = await ExplorerWindowDiscovery.ListenForNewExplorerWindowAsync(currentWindows ?? [],
                cancellationToken: CurrentCancellation);
            EnsureCurrentMerge();
            if (newWindowHandle == 0) return;

            var window = FindTrackedWindowByTopLevel(newWindowHandle);
            if (window == null) return;

            SelectItems(window, windowToOpen.SelectedItems);
        }
        finally
        {
            if (lockToOpenWindows)
                _toOpenWindowsLock.Release();
        }
    }
    private async Task<bool> OpenTabNavigateWithSelection(WindowRecord windowToOpen, nint windowHandle = 0)
    {
        ExplorerDebugLog.Write($"OpenTab begin target={windowToOpen.Location}");
        nint mainWindowHWnd = 0;
        nint newTabHandle = 0;
        WindowIdentity mainWindowIdentity = default;
        WindowIdentity newTabIdentity = default;
        InternetExplorer? window = null;

        await _toOpenWindowsLock.WaitAsync(CurrentCancellation);
        try
        {
            EnsureCurrentMerge();
            ExplorerDebugLog.Write($"OpenTab lock target={windowToOpen.Location}");
            if (_reuseTabs && TrackedWindowCount > 0 && !string.IsNullOrWhiteSpace(windowToOpen.Location))
            {
                if (TrySearchForTab(windowToOpen.Location, windowToOpen.Handle, out var existingTab, out var existingWindow))
                {
                    windowHandle = WinApi.GetParent(existingTab);
                    if (!await SelectTabByHandle(windowHandle, existingTab))
                        return false;
                    EnsureCurrentMerge();
                    if (existingWindow == null || !SelectItems(existingWindow, windowToOpen.SelectedItems))
                    {
                        ExplorerDebugLog.Write($"OpenTab reuse-selection-failed target={windowToOpen.Location}");
                        return false;
                    }
                    ExplorerDebugLog.Write($"OpenTab reused target={windowToOpen.Location}");
                    return true;
                }
            }

            // Get the main window
            mainWindowHWnd = ExplorerWindowDiscovery.IsFileExplorerWindow(windowHandle)
                ? windowHandle
                : GetMainWindowHWnd(windowToOpen.Handle);

            if (mainWindowHWnd == 0)
            {
                if (_currentMerge.Value != null)
                    return false;
                await OpenNewWindowWithSelection(windowToOpen, lockToOpenWindows: false);
                ExplorerDebugLog.Write($"OpenTab opened-window target={windowToOpen.Location}");
                return true;
            }

            try
            {
                mainWindowIdentity = WindowIdentity.Capture(mainWindowHWnd);
                EnsureWindowIdentity(mainWindowIdentity);
                var currentTabs = ExplorerWindowDiscovery.GetAllExplorerTabs(mainWindowHWnd).ToArray();
                ExplorerDebugLog.Write($"OpenTab main={mainWindowHWnd} tabs={currentTabs.Length} target={windowToOpen.Location}");

                await RequestToOpenNewTab(mainWindowHWnd, lockToOpenWindows: false);
                ExplorerDebugLog.Write($"OpenTab requested main={mainWindowHWnd} target={windowToOpen.Location}");

                newTabHandle = await ExplorerWindowDiscovery.ListenForNewExplorerTabAsync(mainWindowHWnd, currentTabs, 2_000,
                    CurrentCancellation);
                EnsureCurrentMerge();
                if (newTabHandle == 0)
                {
                    ExplorerDebugLog.Write($"OpenTab no-new-tab target={windowToOpen.Location}");
                    return false;
                }
                newTabIdentity = WindowIdentity.Capture(newTabHandle);
                EnsureWindowIdentity(mainWindowIdentity);
                EnsureWindowIdentity(newTabIdentity);
                ExplorerDebugLog.Write($"OpenTab new-tab={newTabHandle} target={windowToOpen.Location}");

                window = await Helper.DoUntilNotDefaultAsync(
                    () => FindShellWindowByTabHandle(newTabHandle, mainWindowHWnd),
                    2_000,
                    50, CurrentCancellation);
                EnsureCurrentMerge();

                if (window == null)
                {
                    await CloseFailedNewTabAsync(mainWindowIdentity, newTabIdentity);
                    ExplorerDebugLog.Write($"OpenTab missing-shell-window tab={newTabHandle} target={windowToOpen.Location}");
                    return false;
                }
                ExplorerDebugLog.Write($"OpenTab found-shell-window tab={newTabHandle} target={windowToOpen.Location}");
            }
            catch (Exception ex)
            {
                await CloseFailedNewTabAsync(mainWindowIdentity, newTabIdentity);
                if (IsDisconnectedShell(ex))
                    throw;
                ExplorerDebugLog.Write($"OpenTab error tab={newTabHandle} target={windowToOpen.Location} error={ex.GetType().Name}:{ex.Message}");
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
            EnsureWindowIdentity(mainWindowIdentity);
            EnsureWindowIdentity(newTabIdentity);

            if (!await NavigateNewTabToTargetAsync(window, windowToOpen.Location))
            {
                await CloseFailedNewTabAsync(mainWindowIdentity, newTabIdentity);
                return false;
            }

            EnsureWindowIdentity(mainWindowIdentity);
            EnsureWindowIdentity(newTabIdentity);
            Helper.RestoreWindowToForeground(mainWindowHWnd);
            SelectItems(window, windowToOpen.SelectedItems);
            return true;
        }
        catch (Exception ex)
        {
            await CloseFailedNewTabAsync(mainWindowIdentity, newTabIdentity);
            if (IsDisconnectedShell(ex))
                throw;
            ExplorerDebugLog.Write($"OpenTab post-create error tab={newTabHandle} target={windowToOpen.Location} error={ex.GetType().Name}:{ex.Message}");
            return false;
        }
    }
    private async Task CloseFailedNewTabAsync(WindowIdentity parent, WindowIdentity tab)
    {
        try
        {
            for (var attempt = 0; attempt < 2; attempt++)
            {
                if (!parent.IsCurrent || !tab.IsCurrent || WinApi.GetParent(tab.Handle) != parent.Handle)
                    return;
                EnsureCurrentMerge();
                // Closing a tab destroys its window while the command is handled; post it and observe.
                if (!WinApi.PostMessage(tab.Handle, WinApi.WM_COMMAND, 0xA021, 1) && tab.IsCurrent)
                {
                    ReportStatus("Explorer did not accept cleanup of the newly created tab.");
                    return;
                }
                var exists = await Helper.DoUntilConditionAsync(() => tab.IsCurrent,
                    current => !current, 700, 50, CurrentCancellation);
                if (!exists)
                    return;
            }
            ReportStatus("The newly created tab could not be closed; no further close commands will be sent.");
        }
        catch (OperationCanceledException)
        {
            ExplorerDebugLog.Write("Tab cleanup cancelled; no late close command was sent.");
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
            50, CurrentCancellation);

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
            // Trust the first NavigateComplete2 unconditionally. Explorer can fire the event a few
            // ticks before LocationURL is updated, so gating on AreLocationsEquivalent here would
            // leave tcs unset on the typical fast path and stall every merge until WaitForNavigation
            // catches up.
            navigationCompleted.TrySetResult(true);
        };

        try
        {
            window.NavigateComplete2 += navigateHandler;
            if (!await NavigateToTargetIfNeeded(window, targetLocation))
            {
                ExplorerDebugLog.Write($"OpenTab navigate-failed target={targetLocation}");
                return false;
            }

            ExplorerDebugLog.Write($"OpenTab navigated target={targetLocation}");
            await Task.WhenAny(navigationCompleted.Task, Task.Delay(NavigationCompleteWaitMs, CurrentCancellation));
            EnsureCurrentMerge();

            if (AreLocationsEquivalent(TryGetLocation(window), targetLocation))
                return true;

            if (await WaitForNavigation(window, targetLocation, NavigationVerificationWaitMs))
                return true;

            ExplorerDebugLog.Write($"OpenTab target-check-failed target={targetLocation} current={TryGetLocation(window)}");
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
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            ExplorerDebugLog.Write($"Tab navigation temporarily unavailable error={exception.HResult:X8}");
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

        if (ExplorerWindowDiscovery.IsFileExplorerForeground(out var foregroundWindow) &&
            IsPreferredMergeTargetWindow(foregroundWindow, otherThan, targetLocation))
        {
            _mainWindowHandle = foregroundWindow;
            return _mainWindowHandle;
        }

        if (IsPreferredMergeTargetWindow(_mainWindowHandle, otherThan, targetLocation))
            return _mainWindowHandle;

        var allWindows = _getExplorerWindows().ToArray();
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
    private bool IsStableMergeTargetWindow(nint hWnd, nint otherThan)
    {
        if (hWnd == 0 || hWnd == otherThan)
            return false;
        if (!ExplorerWindowDiscovery.IsShownExplorerWindow(hWnd))
            return false;
        if (!HasHookedShellWindowForTopLevel(hWnd))
            return false;

        return GetActiveTabHandle(hWnd) != 0;
    }
    private static bool IsFallbackMergeTargetWindow(nint hWnd, nint otherThan)
    {
        if (hWnd == 0 || hWnd == otherThan)
            return false;
        if (!ExplorerWindowDiscovery.IsShownExplorerWindow(hWnd))
            return false;

        return GetActiveTabHandle(hWnd) != 0;
    }
    private Task<nint> GetTabHandle(InternetExplorer window)
    {
        if (TryGetKnownTabHandle(window, out var handle))
            return Task.FromResult(handle);

        return QueryTabHandle(window, updateDictionary: true);
    }
    private bool TryPublishTabHandle(InternetExplorer window, nint tabHandle)
    {
        if (tabHandle == 0)
            return false;

        try
        {
            lock (_windowEntryDictLock)
            {
                if (!_windowEntryDict.TryGetValue(window, out WindowInfo? info) ||
                    !IsCurrentWindow(window, info) || WinApi.GetParent(tabHandle) != info.Identity.Handle)
                    return false;

                if (_windowEntryDict.TryGetValue(tabHandle, out InternetExplorer? previousWindow) &&
                    previousWindow != null && !ReferenceEquals(previousWindow, window))
                {
                    if (!_windowEntryDict.TryGetValue(previousWindow, out WindowInfo? previousInfo))
                        return false;
                    if (IsCurrentWindow(previousWindow, previousInfo) && IsCurrentTab(previousInfo, tabHandle))
                    {
                        RemoveWindowAndUnhookEvents(window, info, restoreHiddenWindow: false);
                        ExplorerDebugLog.Write($"Registered duplicate tab={tabHandle} hwnd={info.Identity.Handle}");
                        return true;
                    }
                    RemoveWindowAndUnhookEvents(previousWindow, previousInfo);
                }

                var tabIdentity = WindowIdentity.Capture(tabHandle);
                if (!tabIdentity.IsCurrent)
                    return false;
                _windowEntryDict.UpdateOptionalKey(window, tabHandle);
                info.TabIdentity = tabIdentity;
                return true;
            }
        }
        catch (Exception ex) when (!IsDisconnectedShell(ex))
        {
            ExplorerDebugLog.Write($"Tab handle publish failed handle={tabHandle} error={ex.GetType().Name}:{ex.Message}");
            return false;
        }
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
                EnsureCurrentMerge();

                if (updateDictionary && hWnd != 0 && !TryPublishTabHandle(window, hWnd))
                    return 0;

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
        if (!_staTaskScheduler.IsCurrentThread)
            return await Task.Factory.StartNew(() => FindShellWindowByTabHandle(tabHandle, parentWindowHandle),
                CurrentCancellation, TaskCreationOptions.DenyChildAttach, _staTaskScheduler).Unwrap();

        EnsureCurrentMerge();
        var cachedWindow = GetWindowByTabHandle(tabHandle, parentWindowHandle);
        if (cachedWindow != null)
            return cachedWindow;

        var count = _shellWindows.Count;
        for (var i = count - 1; i >= 0; i--)
        {
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;

            if (parentWindowHandle != 0 && SafeGetWindowHandle(window) != parentWindowHandle)
                continue;

            var currentTabHandle = await QueryTabHandle(window, updateDictionary: false);
            EnsureCurrentMerge();
            if (currentTabHandle != tabHandle)
                continue;

            WindowInfo windowInfo;
            InternetExplorer windowToReturn = window;
            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.TryGetValue(tabHandle, out InternetExplorer? existingWindow) && existingWindow != null &&
                    _windowEntryDict.TryGetValue(existingWindow, out windowInfo!) &&
                    IsCurrentWindow(existingWindow, windowInfo) && IsCurrentTab(windowInfo, tabHandle))
                {
                    windowToReturn = existingWindow;
                }
                else
                {
                    if (!_windowEntryDict.TryGetValue(window, out windowInfo!))
                    {
                        windowInfo = CreateWindowInfo(window);
                        _windowEntryDict.Add(window, windowInfo);
                    }
                    if (!TryPublishTabHandle(window, tabHandle))
                        continue;
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
    private int TrackedWindowCount
    {
        get
        {
            lock (_windowEntryDictLock)
                return _windowEntryDict.Count;
        }
    }
    private bool TryGetTrackedEntry(InternetExplorer window, out WindowEntry entry)
    {
        lock (_windowEntryDictLock)
            return _windowEntryDict.TryGetValue(window, out entry);
    }
    private InternetExplorer? GetWindowByTabHandle(nint tabHandle, nint parentWindowHandle)
    {
        if (tabHandle == 0) return null;

        InternetExplorer? window;
        lock (_windowEntryDictLock)
        {
            if (!_windowEntryDict.TryGetValue(tabHandle, out window) || window == null ||
                !_windowEntryDict.TryGetValue(window, out WindowInfo? info) ||
                !IsCurrentWindow(window, info) || !IsCurrentTab(info, tabHandle))
                return null;

            return parentWindowHandle == 0 || info.Identity.Handle == parentWindowHandle ? window : null;
        }
    }
    private static string[]? GetSelectedItems(InternetExplorer window)
    {
        if (window.Document is not ShellFolderView document)
            return null;
        var selectedItems = document.SelectedItems();
        if (selectedItems == null)
            return null;
        var count = selectedItems.Count;
        if (count == 0) return Array.Empty<string>();

        var result = new string[count];
        for (var index = 0; index < count; index++)
        {
            var item = selectedItems.Item(index);
            if (item == null)
                return null;
            result[index] = item.Name;
        }

        return result;
    }
    private static string[]? TryGetSelectedItems(InternetExplorer window)
    {
        try
        {
            return GetSelectedItems(window);
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            ExplorerDebugLog.Write($"Selection capture unavailable error={exception.GetType().Name}");
            return null;
        }
    }
    private bool SelectItems(InternetExplorer window, string[]? names)
    {
        if (names == null || names.Length == 0) return true;
        var identity = WindowIdentity.Capture(SafeGetWindowHandle(window));
        EnsureWindowIdentity(identity);

        if (window.Document is not ShellFolderView document) return false;

        const int selectItem = 1;
        const int deselectOthers = 4;
        var selectedAny = false;
        foreach (var name in names)
        {
            EnsureWindowIdentity(identity);
            object item = document.Folder.ParseName(name);
            if (item == null) continue;
            document.SelectItem(ref item, selectedAny ? selectItem : selectItem | deselectOthers);
            selectedAny = true;
        }
        return selectedAny;
    }
    private static string GetLocation(InternetExplorer window)
    {
        var path = window.LocationURL;
        if (!string.IsNullOrWhiteSpace(path)) return Helper.NormalizeLocation(path);

        // Recycle Bin, This PC, etc
        if (window.Document is not ShellFolderView document || document.Folder is not Folder2 folder)
            return string.Empty;
        return Helper.NormalizeLocation(folder.Self.Path);
    }
    private async Task Navigate(InternetExplorer window, string path)
    {
        var identity = WindowIdentity.Capture(SafeGetWindowHandle(window));
        EnsureWindowIdentity(identity);
        if (!path.Contains('#') && !path.Contains("%23"))
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
            EnsureWindowIdentity(identity);
            window.Navigate2(folder);
        }
        finally
        {
            if (folder != null)
                Marshal.ReleaseComObject(folder);
        }
    }
    private Task RunInStaThread(Action action, TaskCreationOptions options = default, CancellationToken cancellationToken = default)
    {
        return RunInStaThread(() =>
        {
            action();
            return true;
        }, options, cancellationToken);
    }

    private Task<T?> RunInStaThread<T>(Func<T?> action, TaskCreationOptions options = default, CancellationToken cancellationToken = default)
    {
        var generation = _shellGeneration;
        var token = cancellationToken.CanBeCanceled ? cancellationToken : CurrentCancellation;
        T? Execute()
        {
            token.ThrowIfCancellationRequested();
            EnsureCurrentMerge();
            if (generation != _shellGeneration)
                throw new OperationCanceledException("The Explorer connection has changed.");
            return action();
        }
        return _staTaskScheduler.IsCurrentThread
            ? Task.FromResult(Execute())
            : Task.Factory.StartNew(Execute, token, options, _staTaskScheduler);
    }

    private void StartExplorerProcessCheck() => _explorerCheckTimer = new Timer(CheckForMainExplorer, null, 0, 1000);

    private void CheckForMainExplorer(object? state)
    {
        if (_disposed)
            return;
        if (_mainExplorerProcessId != 0)
        {
            if (HasPendingTabRegistrations() || HasUnregisteredExplorerTabs())
                ScheduleShellWindowRegistration();
            return;
        }
        if (Interlocked.CompareExchange(ref _shellTransitionScheduled, 1, 0) != 0)
            return;
        using var process = ExplorerWindowDiscovery.GetMainExplorerProcess();
        if (process == null)
        {
            Volatile.Write(ref _shellTransitionScheduled, 0);
            return;
        }
        _shellTransitionTask = InitializeShellAsync(process.Id);
    }

    private void OnExplorerProcessTerminated(object? sender, ProcessEventArgs eventArgs)
    {
        if (_disposed)
            return;
        if (eventArgs.ProcessId == _mainExplorerProcessId)
        {
            RetireShellConnection("explorer-restarted");
        }
        else
        {
            ScheduleShellWindowRegistration();
        }
    }

    private async Task InitializeShellObjectsAsync()
    {
        _preExistingExplorerWindowsProtected = false;
        _shellPathComparer = new ShellPathComparer();
        _shellWindows = new ShellWindows();
        _shellLifetime.Token.ThrowIfCancellationRequested();

        _defaultLocation = GetDefaultExplorerLocation();
        ClearShellCaches();
        RecoverHiddenExplorerWindows("initialize-shell");

        if (ExplorerWindowDiscovery.IsFileExplorerForeground(out var foregroundWindow))
            _mainWindowHandle = foregroundWindow;

        // Hook the global "WindowRegistered" event
        _windowRegisteredHandler = OnShellWindowRegistered;
        _shellWindows.WindowRegistered += _windowRegisteredHandler;

        // WinEvent only wakes ShellWindows processing; WindowRegistered owns merge and release.
        _shellLifetime.Token.ThrowIfCancellationRequested();
        _eventObjectShowHookCallback = OnWindowShown;
        _winEventHookThread = new WinEventHookThread(_eventObjectShowHookCallback);
        _winEventHookThread.Start();

        // Hook the event handlers for already-open windows
        var registrations = new List<Task>();
        var count = _shellWindows.Count;
        for (var i = 0; i < count; i++)
        {
            _shellLifetime.Token.ThrowIfCancellationRequested();
            if (_shellWindows.Item(i) is not InternetExplorer window)
                continue;

            WindowInfo windowInfo;
            lock (_windowEntryDictLock)
            {
                if (_windowEntryDict.Keys.Contains(window))
                    continue;

                windowInfo = CreateWindowInfo(window);
                _windowEntryDict.Add(window, windowInfo);

            }

            PreventWindowHiding(new IntPtr(window.HWND));

            if (_mainWindowHandle == 0)
                _mainWindowHandle = new IntPtr(window.HWND);

            registrations.Add(RegisterIndependentWindowAsync(window, windowInfo, windowInfo.Identity.Handle));
        }

        await Task.WhenAll(registrations);
        _shellLifetime.Token.ThrowIfCancellationRequested();
        _preExistingExplorerWindowsProtected = true;
        ScheduleShellWindowRegistration();
        ConcealPreloadedExplorerFrames();
        ObserveDesktopFolderOpen();
        if (_isForcingTabs)
            StartMergeSourceConcealPulse(500);
    }

    private string GetDefaultExplorerLocation()
    {
        var location = _getDefaultExplorerLaunchId() switch
        {
            2 => "shell:::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}",
            3 => "shell:::{088E3905-0323-4B02-9826-5D99428E115F}",
            4 => "shell:::{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
            _ => "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        };

        var pidl = _shellPathComparer.GetPidlFromPath(location);
        var path = ShellPathComparer.GetPathFromPidl(pidl);
        Marshal.FreeCoTaskMem(pidl);

        return Helper.NormalizeLocation(path ?? location);
    }

    private void DisposeShellObjects()
    {
        _preExistingExplorerWindowsProtected = false;
        StopMergeSourceConcealPulse();
        ReleaseDesktopFolderOpenObserver();
        RecoverHiddenExplorerWindows("dispose-shell");
        var hookThread = Interlocked.Exchange(ref _winEventHookThread, null);
        hookThread?.Dispose();
        _eventObjectShowHookCallback = null;

        List<WindowEntry> windowEntries;
        lock (_windowEntryDictLock)
        {
            windowEntries = ((IEnumerable<WindowEntry>)_windowEntryDict).ToList();
            _windowEntryDict.Clear();
        }
        foreach (var (window, info) in windowEntries)
        {
            try
            {
                if (info.OnQuitHandler != null) window.OnQuit -= info.OnQuitHandler;
                if (info.OnNavigateHandler != null) window.NavigateComplete2 -= info.OnNavigateHandler;
                Marshal.ReleaseComObject(window);
                info.Identity.Release();
            }
            catch (Exception exception) when (exception is COMException or InvalidComObjectException)
            {
                ExplorerDebugLog.Write($"Shell window cleanup failed error={exception.GetType().Name}");
            }
        }

        if (_shellWindows != null)
        {
            try
            {
                if (_windowRegisteredHandler != null)
                    _shellWindows.WindowRegistered -= _windowRegisteredHandler;
                Marshal.ReleaseComObject(_shellWindows);
            }
            catch (Exception exception) when (exception is COMException or InvalidComObjectException)
            {
                ExplorerDebugLog.Write($"Shell connection cleanup failed error={exception.GetType().Name}");
            }
        }
        _windowRegisteredHandler = null;
        _shellWindows = null!;
        _shellPathComparer?.Dispose();
        _shellPathComparer = null!;
        lock (_closedWindowsLock)
            _closedWindows.Clear();
        ClearShellCaches();
        _processedHWnds.Clear();
        _mainWindowHandle = 0;
    }

    private void ClearShellCaches()
    {
        _registrationRetryPending = false;
        _startupLocationCache.Clear();
        TabStrip.Clear();
        _hookedTopLevelUseCounts.Clear();
    }

    private void RecoverHiddenExplorerWindows(string reason)
    {
        var restored = 0;
        foreach (var concealed in _mergeSourceHWnds.Values.ToArray())
        {
            if (RestoreConcealedWindow(concealed))
                restored++;
        }
        restored += ExplorerWindowVisibility.RestoreAll();
        _closingMergeSourceHWnds.Clear();
        if (restored > 0)
            ExplorerDebugLog.Write($"RecoverHiddenExplorerWindows reason={reason} restored={restored}");
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        var startedAt = Environment.TickCount64;
        _disposed = true;
        _isForcingTabs = false;
        _preExistingExplorerWindowsProtected = false;
        Interlocked.Increment(ref _hookGeneration);
        Interlocked.Increment(ref _shellGeneration);
        _shellLifetime.Cancel();
        _explorerCheckTimer?.Dispose();
        _hookLifetime.Cancel();
        _mergeSafetyTimer.Dispose();
        _selectionTimer.Dispose();
        StopMergeSourceConcealPulse();
        Interlocked.Exchange(ref _winEventHookThread, null)?.Dispose();
        RecoverHiddenExplorerWindows("dispose");
        TabStrip.Dispose();
        _processWatcher.ProcessTerminated -= OnExplorerProcessTerminated;
        _processWatcher.Dispose();

        var cleanup = Task.Run(FinishDisposalAsync);
        var remaining = Math.Max(0, 2_000 - (int)(Environment.TickCount64 - startedAt));
        if (!cleanup.Wait(remaining))
            ExplorerDebugLog.Write("Shell cleanup is still pending; shutdown will not wait longer.");
        GC.SuppressFinalize(this);
    }
}
