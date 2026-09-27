using SHDocVw;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Models;

namespace WinTab.Hooks;

using WindowEntry = DualKeyEntry<InternetExplorer, nint?, WindowInfo>;

/// <summary>
/// Adopting the windows ShellWindows registers: deciding between merging and keeping them, and attaching
/// to or detaching from their browser events.
/// </summary>
public partial class ExplorerWatcher
{
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
                    // ShellWindows can publish a preload before the shell shows it. Registering it as
                    // an independent window would remove its concealment and protect a later launch.
                    if (ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd) && !WinAPI.WinApi.IsWindowVisible(hWnd))
                        continue;
                    wasTrackedTopLevel = HasTrackedTopLevelWindow(hWnd);
                    var tabCount = ExplorerWindowDiscovery.GetAllExplorerTabs(hWnd).Take(2).Count();
                    if (tabCount <= 1 &&
                        (wasTrackedTopLevel || !singleTabTopLevelsInBatch.Add(hWnd)))
                        continue;

                    if (!wasTrackedTopLevel &&
                        !IsWindowProtected(hWnd) &&
                        _isForcingTabs &&
                        !IsIndependentOpenRequested() &&
                        MainWindowHandle != hWnd &&
                        !ReleaseTornOffTabWindow(hWnd))
                    {
                        HideMergeSourceWindow(hWnd);
                    }

                    _windowEntryDict.Add(window, windowInfo);
                    if (_windowEntryDict.Count == 1)
                        MainWindowHandle = hWnd;
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

    private void ScheduleShellWindowRegistration()
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
        if (await TryRestoreNewExplorerWindowAsync(window, windowInfo))
            return;
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

            if (IsIndependentOpenRequested())
            {
                ExplorerDebugLog.Write($"Registered release ctrl-shift hwnd={hWnd}");
                await RestoreMergeSourceWindowAsync(hWnd);
                await RegisterIndependentWindowAsync(window, windowInfo, hWnd);
                return;
            }

            // The window is shown again and registered as an independent window on the way out.
            if (ReleaseTornOffTabWindow(hWnd))
            {
                ExplorerDebugLog.Write($"Registered release torn-off-tab hwnd={hWnd}");
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

            if (await WaitForTornOffTabWindowAsync(hWnd))
            {
                ExplorerDebugLog.Write($"Registered release torn-off-tab-shown hwnd={hWnd} location={location}");
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
        if (_captureSessions)
            CaptureExplorerSessions();
    }

    private void HookWindowEvents(InternetExplorer window, WindowInfo windowInfo)
    {
        if (windowInfo.EventsHooked || !IsCurrentWindow(window, windowInfo))
            return;

        var hookedTopLevelHWnd = SafeGetWindowHandle(window);
        if (hookedTopLevelHWnd != 0)
        {
            windowInfo.HookedTopLevelHWnd = hookedTopLevelHWnd;
            _hookedTopLevelUseCounts.AddOrUpdate(windowInfo.Identity, 1, (_, count) => count + 1);
        }

        // Create a strongly-typed handler so we can remove it later
        windowInfo.OnQuitHandler = () =>
        {
            if (windowInfo.Closed || !IsRegisteredWindow(window, windowInfo))
                return;
            // Explorer waits for this callback. Retire the tab immediately, but do not call back into
            // its dying view or connection point while its UI thread is waiting for our response.
            windowInfo.Closed = true;
            if (_captureSessions)
            {
                _sessionTracker.TabClosing(windowInfo.Identity, windowInfo.TabIdentity, Environment.TickCount64);
                RecordClosedTab(windowInfo);
                _selectionWork.Request();
            }
            // Use the selection captured while the view was alive, never read Document during OnQuit.
            // Home, This PC, etc. carry nothing worth restoring.
            var location = windowInfo.Location;
            if (!string.IsNullOrWhiteSpace(location) && location != _defaultLocation)
                RememberClosedWindow(new WindowRecord(location, windowInfo.Identity.Handle, windowInfo.SelectedItems));

            // The coalesced registration worker removes closed records on the owning STA after we
            // return. A burst of closes shares one queued pass, rather than one COM task per tab.
            ScheduleShellWindowRegistration();
        };
        windowInfo.OnNavigateHandler = (object _, ref object url) =>
        {
            if (!IsCurrentWindow(window, windowInfo))
                return;
            var location = url?.ToString();
            if (_captureSessions && _restoringSessionWindows.TryGetValue(windowInfo.Identity, out var attempt) &&
                attempt.InitialTab == windowInfo.TabIdentity)
                attempt.InitialTabNavigated();
            if (!string.IsNullOrWhiteSpace(location))
            {
                windowInfo.Location = Helper.NormalizeLocation(location);
                if (_captureSessions)
                    _sessionTracker.UpdateLocation(windowInfo.Identity, windowInfo.TabIdentity, windowInfo.Location);
            }
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

        var identity = windowInfo.Identity;
        var remaining = _hookedTopLevelUseCounts.AddOrUpdate(identity, 0, (_, count) => Math.Max(0, count - 1));
        if (remaining == 0)
            _hookedTopLevelUseCounts.TryRemove(new KeyValuePair<WindowIdentity, int>(identity, 0));

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
            if (MainWindowHandle == hWnd && !ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd))
                MainWindowHandle = 0;
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
}
