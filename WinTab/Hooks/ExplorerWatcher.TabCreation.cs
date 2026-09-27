using Shell32;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Models;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Opening a location as a new tab of a window, or as a window of its own, and cleaning up after a failure.
/// </summary>
public partial class ExplorerWatcher
{
    /// <summary>
    /// SBSP_NEWBROWSER combined with the new-tab flag Explorer itself uses for a navigation-pane middle click.
    /// Explorer then appends a background tab that is created already at the requested location, so the
    /// window keeps showing its current tab and no default page appears first. Where the tab flag is unknown,
    /// SBSP_NEWBROWSER makes Explorer open a window instead of navigating the caller's tab away; that
    /// window is detected and the classic new-tab command is used from then on.
    /// </summary>
    private const uint NewTabBrowseFlags = 0x0402;
    private const int NewTabWaitMs = 2_000;
    /// <summary>
    /// Windows 11 preloads a hidden Explorer frame at any time. A new window therefore only answers a direct
    /// tab request when no tab has followed it within this time.
    /// </summary>
    private const int WindowAnswerGraceMs = 250;
    private const int AppendedTabActivationWaitMs = 1_500;
    /// <summary>How long the active tab is watched after the frame is brought to the front.</summary>
    private const int ForegroundSettleWaitMs = 250;
    /// <summary>Budget for closing a tab a failed merge created; independent of the merge's own deadline.</summary>
    private const int FailedTabCleanupTimeoutMs = 2_000;
    private const int ClosedWindowHistoryLimit = 100;

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
        // Set when Explorer created the tab at the location itself; such a tab is appended in the background
        // at this index and is brought to the front once it has arrived at the location. Until it is in
        // front, focus Explorer reports for its view is this merge's doing, not a native file-location request.
        var createdAtLocation = false;
        var appendedIndex = 0;
        var focusGuarded = false;
        var lockHeld = false;

        try
        {
            await _toOpenWindowsLock.WaitAsync(CurrentCancellation);
            lockHeld = true;
            // Keep creation and activation in one critical section. A second merge must not foreground
            // this frame while the first merge still has a newly appended, unselected tab settling.
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
                    if (!await PrepareWindowForTabCreationAsync(mainWindowIdentity))
                        return false;
                    var currentTabs = ExplorerWindowDiscovery.GetAllExplorerTabs(mainWindowHWnd).ToArray();
                    ExplorerDebugLog.Write($"OpenTab main={mainWindowHWnd} tabs={currentTabs.Length} target={windowToOpen.Location}");

                    appendedIndex = currentTabs.Length;
                    newTabHandle = await CreateTabAtLocationAsync(mainWindowHWnd, mainWindowIdentity, currentTabs, windowToOpen.Location);
                    createdAtLocation = newTabHandle != 0;
                    if (createdAtLocation)
                    {
                        Interlocked.Increment(ref _tabSelectionsInProgress);
                        focusGuarded = true;
                    }
                    else
                    {
                        await RequestToOpenNewTab(mainWindowHWnd, lockToOpenWindows: false);
                        ExplorerDebugLog.Write($"OpenTab requested main={mainWindowHWnd} target={windowToOpen.Location}");

                        newTabHandle = await ExplorerWindowDiscovery.ListenForNewExplorerTabAsync(mainWindowHWnd, currentTabs, NewTabWaitMs,
                            CurrentCancellation);
                        EnsureCurrentMerge();
                        if (newTabHandle == 0)
                        {
                            ExplorerDebugLog.Write($"OpenTab no-new-tab target={windowToOpen.Location}");
                            return false;
                        }
                    }
                    var candidateIdentity = WindowIdentity.Capture(newTabHandle);
                    // A direct request can race with a user-created tab. It is not ours to clean up until
                    // its location confirms the request; classic tabs still need explicit navigation.
                    if (!createdAtLocation)
                        newTabIdentity = candidateIdentity;
                    EnsureWindowIdentity(mainWindowIdentity);
                    EnsureWindowIdentity(candidateIdentity);
                    ExplorerDebugLog.Write($"OpenTab new-tab={newTabHandle} target={windowToOpen.Location}");

                    window = await Helper.DoUntilNotDefaultAsync(
                        () => FindShellWindowByTabHandle(newTabHandle, mainWindowHWnd),
                        2_000,
                        50, CurrentCancellation);
                    EnsureCurrentMerge();

                    if (createdAtLocation)
                    {
                        if (window == null || !await WaitForNavigation(window, windowToOpen.Location, NavigationVerificationWaitMs))
                        {
                            ExplorerDebugLog.Write($"OpenTab direct ownership unconfirmed tab={newTabHandle} target={windowToOpen.Location}");
                            return false;
                        }
                        EnsureWindowIdentity(mainWindowIdentity);
                        EnsureWindowIdentity(candidateIdentity);
                        newTabIdentity = candidateIdentity;
                    }
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

            try
            {
                if (window == null)
                    return false;
                EnsureWindowIdentity(mainWindowIdentity);
                EnsureWindowIdentity(newTabIdentity);

                if (!await NavigateNewTabToTargetAsync(window, windowToOpen.Location, navigationStarted: createdAtLocation))
                {
                    await CloseFailedNewTabAsync(mainWindowIdentity, newTabIdentity);
                    return false;
                }

                EnsureWindowIdentity(mainWindowIdentity);
                EnsureWindowIdentity(newTabIdentity);
                if (createdAtLocation &&
                    !await ActivateAppendedTabAsync(mainWindowHWnd, mainWindowIdentity, newTabHandle, newTabIdentity, appendedIndex))
                {
                    ExplorerDebugLog.Write($"OpenTab activation-failed tab={newTabHandle} target={windowToOpen.Location}");
                    await CloseFailedNewTabAsync(mainWindowIdentity, newTabIdentity);
                    return false;
                }
                if (!createdAtLocation)
                    Helper.RestoreWindowToForeground(mainWindowHWnd);
                if (!SelectItems(window, windowToOpen.SelectedItems))
                {
                    await CloseFailedNewTabAsync(mainWindowIdentity, newTabIdentity);
                    ExplorerDebugLog.Write($"OpenTab selection-failed tab={newTabHandle} target={windowToOpen.Location}");
                    return false;
                }
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
        finally
        {
            if (focusGuarded)
            {
                Volatile.Write(ref _ignoreNativeFocusThrough, Environment.TickCount);
                Interlocked.Decrement(ref _tabSelectionsInProgress);
            }
            if (lockHeld)
                _toOpenWindowsLock.Release();
        }
    }

    /// <summary>
    /// BrowseObject itself foregrounds the frame and then refocuses its old view, even for a background tab.
    /// Activate the settled frame once BEFORE creating anything so that native focus transfer is a no-op.
    /// Once a new tab exists, activation still strictly selects the tab before any further foreground work.
    /// The caller holds the merge lock through creation and selection to keep that boundary safe.
    /// </summary>
    private async Task<bool> PrepareWindowForTabCreationAsync(WindowIdentity identity)
    {
        EnsureWindowIdentity(identity);
        Helper.RestoreWindowToForeground(identity.Handle);
        var foreground = await Helper.DoUntilConditionAsync(ExplorerNavigationAccess.ForegroundFrame,
            handle => handle == identity.Handle, 500, 10, CurrentCancellation);
        EnsureWindowIdentity(identity);
        if (foreground == identity.Handle)
            return true;
        ExplorerDebugLog.Write($"Tab creation cancelled; target did not reach foreground hwnd={identity.Handle}");
        return false;
    }

    /// <summary>
    /// Asks Explorer, through the browser of one of the target window's tabs, to open the location as a new tab
    /// of that window. Explorer appends the tab in the background and creates it at the location, so the window
    /// keeps showing its current tab and no default page is shown first. Returns 0 when the classic new-tab
    /// command has to be used: nothing is tracked for the window, the location has no shell item, or Explorer
    /// answered with a window or with nothing.
    /// </summary>
    private async Task<nint> CreateTabAtLocationAsync(nint mainWindowHWnd, WindowIdentity mainWindowIdentity,
        nint[] currentTabs, string location)
    {
        if (_directTabUnsupported || string.IsNullOrWhiteSpace(location))
            return 0;
        // The active tab's browser is the one least likely to be closed while the request is under way.
        var browser = GetWindowByTabHandle(GetActiveTabHandle(mainWindowHWnd), mainWindowHWnd) ??
            FindTrackedWindowByTopLevel(mainWindowHWnd, (_, info) => info.EventsHooked && !info.Closed);
        if (browser == null)
        {
            ExplorerDebugLog.Write($"OpenTab direct skipped; no tracked tab main={mainWindowHWnd}");
            return 0;
        }

        var knownWindows = new HashSet<nint>(_getExplorerWindows());
        if (!await RunInStaThread(() => RequestTabAtLocation(browser, mainWindowIdentity, location)))
            return 0;

        return await WaitForTabAtLocationAsync(mainWindowHWnd, new HashSet<nint>(currentTabs), knownWindows, location);
    }

    /// <summary>
    /// The tab Explorer appends for a direct request, or 0 when Explorer answered with a top-level window or
    /// with nothing at all. Both answers mean the tab flag is not understood: the classic command is used from
    /// then on, so at most one merge pays for finding out.
    /// </summary>
    private async Task<nint> WaitForTabAtLocationAsync(nint mainWindowHWnd, HashSet<nint> knownTabs, HashSet<nint> knownWindows,
        string location, int timeoutMs = NewTabWaitMs)
    {
        long windowSeenAt = 0;
        var outcome = await Helper.DoUntilConditionAsync(() =>
            {
                var tab = ExplorerWindowDiscovery.GetUniqueNewExplorerTab(mainWindowHWnd, knownTabs);
                // Windows 11 preloads hidden frames at any time; only a frame Explorer has shown is its
                // answer to the request. A hidden one is not, and the tab may still arrive within the wait.
                if (tab != 0 || !_getExplorerWindows().Any(handle => !knownWindows.Contains(handle) && WinApi.IsWindowVisible(handle)))
                    return (Tab: tab, WindowOpened: false);
                if (windowSeenAt == 0)
                    windowSeenAt = Environment.TickCount64;
                return (Tab: 0, WindowOpened: Environment.TickCount64 - windowSeenAt >= WindowAnswerGraceMs);
            },
            result => result.Tab != 0 || result.WindowOpened, timeoutMs, 20, CurrentCancellation);
        EnsureCurrentMerge();
        if (outcome.Tab != 0)
        {
            ExplorerDebugLog.Write($"OpenTab direct tab={outcome.Tab} main={mainWindowHWnd} target={location}");
            return outcome.Tab;
        }

        _directTabUnsupported = true;
        ExplorerDebugLog.Write($"OpenTab direct unsupported window-opened={outcome.WindowOpened} main={mainWindowHWnd} target={location}");
        ReportStatus(outcome.WindowOpened
            ? "Explorer opened a window instead of a tab at the requested location; new tabs now use the classic command."
            : "Explorer did not create a tab at the requested location; new tabs now use the classic command.");
        return 0;
    }

    /// <summary>Issues the new-tab request on the shell thread; true when Explorer accepted it and the tab must be observed.</summary>
    private bool RequestTabAtLocation(InternetExplorer browser, WindowIdentity mainWindowIdentity, string location)
    {
        EnsureWindowIdentity(mainWindowIdentity);
        var startedAt = System.Diagnostics.Stopwatch.GetTimestamp();
        void LogStage(string stage) => ExplorerDebugLog.Write(
            $"OpenTab direct request stage={stage} main={mainWindowIdentity.Handle} target={location} elapsedMs={System.Diagnostics.Stopwatch.GetElapsedTime(startedAt).TotalMilliseconds:F0}");
        LogStage("service-provider");
        // ReSharper disable once SuspiciousTypeConversion.Global
        if (browser is not Interop.IServiceProvider serviceProvider)
            return false;

        LogStage("resolve-pidl");
        var pidl = _shellPathComparer.GetPidlFromPath(location);
        if (pidl == 0)
        {
            ExplorerDebugLog.Write($"OpenTab direct skipped; no shell item target={location}");
            return false;
        }

        try
        {
            LogStage("query-browser");
            serviceProvider.QueryService(ref _shellBrowserGuid, ref _shellBrowserGuid, out var shellBrowser);
            if (shellBrowser == null)
                return false;
            try
            {
                EnsureWindowIdentity(mainWindowIdentity);
                LogStage("browse-object");
                var result = shellBrowser.BrowseObject(pidl, NewTabBrowseFlags);
                LogStage($"browse-returned-{result:X8}");
                if (result == 0)
                    return true;
                ExplorerDebugLog.Write($"OpenTab direct rejected hr={result:X8} target={location}");
                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(shellBrowser);
            }
        }
        catch (COMException exception) when (!IsDisconnectedShell(exception))
        {
            ExplorerDebugLog.Write($"OpenTab direct unavailable error={exception.HResult:X8} target={location}");
            return false;
        }
        finally
        {
            Marshal.FreeCoTaskMem(pidl);
            LogStage("finished");
        }
    }

    /// <summary>
    /// Brings a tab Explorer appended in the background to the front. Its index is known because Explorer
    /// appends new tabs, but the outcome is verified; if the index no longer matches, the tab is found by
    /// cycling as for a reused tab.
    /// </summary>
    private async Task<bool> ActivateAppendedTabAsync(nint windowHandle, WindowIdentity windowIdentity,
        nint tabHandle, WindowIdentity tabIdentity, int tabIndex)
    {
        Interlocked.Increment(ref _tabSelectionsInProgress);
        try
        {
            EnsureWindowIdentity(windowIdentity);
            EnsureWindowIdentity(tabIdentity);
            // The frame is not brought to the foreground before the switch. Explorer's tab strip is still
            // settling the tab it appended in the background, and a foreground change at that moment has
            // crashed Explorer (a C++ exception in Windows.UI.FileExplorer on the focus event). The frame is
            // only un-minimized so the command is handled, and is brought to the front once the switch is done.
            if (WinApi.IsIconic(windowHandle))
                WinApi.ShowWindow(windowHandle, WinApi.SW_SHOWNOACTIVATE);
            if (GetActiveTabHandle(windowHandle) == tabHandle)
                return await BringActivatedTabToFrontAsync(windowHandle, windowIdentity, tabHandle, tabIdentity);

            var previousTab = GetActiveTabHandle(windowHandle);
            SelectTabByIndex(windowHandle, tabIndex);
            // Explorer has handled the command once the active tab changes; a change to another tab means
            // the index no longer described the appended tab.
            var active = await Helper.DoUntilConditionAsync(() => GetActiveTabHandle(windowHandle),
                handle => handle == tabHandle || (handle != previousTab && handle != 0),
                AppendedTabActivationWaitMs, 20, CurrentCancellation);
            EnsureWindowIdentity(windowIdentity);
            EnsureWindowIdentity(tabIdentity);
            if (active == tabHandle)
                return await BringActivatedTabToFrontAsync(windowHandle, windowIdentity, tabHandle, tabIdentity);

            ExplorerDebugLog.Write($"Appended tab not active by index hwnd={windowHandle} tab={tabHandle} index={tabIndex} active={active}");
            if (!await SelectTabByHandle(windowHandle, tabHandle, bringToFront: false))
                return false;
            return await BringActivatedTabToFrontAsync(windowHandle, windowIdentity, tabHandle, tabIdentity);
        }
        finally
        {
            // Explorer reports the new view's focus while the tab is activated; that is this activation, not
            // a native file-location request.
            Volatile.Write(ref _ignoreNativeFocusThrough, Environment.TickCount);
            Interlocked.Decrement(ref _tabSelectionsInProgress);
        }
    }
    /// <summary>
    /// Brings the frame to the front after its appended tab has been activated. Restoring a frame can make
    /// Explorer refocus the view that was active before; the active tab is watched for a moment and the
    /// appended tab is selected again if the foreground change undid the switch.
    /// </summary>
    private async Task<bool> BringActivatedTabToFrontAsync(nint windowHandle, WindowIdentity windowIdentity,
        nint tabHandle, WindowIdentity tabIdentity)
    {
        ExplorerDebugLog.Write($"Appended tab foreground hwnd={windowHandle} tab={tabHandle} foreground={WinApi.GetForegroundWindow()}");
        Helper.RestoreWindowToForeground(windowHandle);
        var active = await Helper.DoUntilConditionAsync(() => GetActiveTabHandle(windowHandle),
            handle => handle != tabHandle, ForegroundSettleWaitMs, 20, CurrentCancellation);
        EnsureWindowIdentity(windowIdentity);
        EnsureWindowIdentity(tabIdentity);
        if (active == tabHandle)
            return true;

        ExplorerDebugLog.Write($"Appended tab lost to foreground change hwnd={windowHandle} tab={tabHandle} active={active}");
        return await SelectTabByHandle(windowHandle, tabHandle, bringToFront: false);
    }

    /// <summary>
    /// Closes a tab this merge created but could not finish with. The tab is known to be ours, so the cleanup
    /// gets its own bounded budget instead of the merge's: a merge that ran out of time while navigating or
    /// waiting for the activation lock would otherwise leave the tab behind for good. Stopping the hook or
    /// losing the shell still cancels it, and the window identities are checked before every command.
    /// </summary>
    private async Task CloseFailedNewTabAsync(WindowIdentity parent, WindowIdentity tab)
    {
        if (_disposed || tab.Handle == 0)
            return;
        using var budget = CancellationTokenSource.CreateLinkedTokenSource(_shellLifetime.Token, _hookLifetime.Token);
        budget.CancelAfter(FailedTabCleanupTimeoutMs);
        try
        {
            for (var attempt = 0; attempt < 2; attempt++)
            {
                if (!parent.IsCurrent || !tab.IsCurrent || WinApi.GetParent(tab.Handle) != parent.Handle)
                    return;
                budget.Token.ThrowIfCancellationRequested();
                _internalTabCloses[tab] = 0;
                _closedTabs.Ignore(tab);
                // Closing a tab destroys its window while the command is handled; post it and observe.
                if (!WinApi.PostMessage(tab.Handle, WinApi.WM_COMMAND, 0xA021, 1) && tab.IsCurrent)
                {
                    ReportStatus("Explorer did not accept cleanup of the newly created tab.");
                    return;
                }
                var exists = await Helper.DoUntilConditionAsync(() => tab.IsCurrent,
                    current => !current, 700, 50, budget.Token);
                if (!exists)
                    return;
            }
            ReportStatus("The newly created tab could not be closed; no further close commands will be sent.");
        }
        catch (OperationCanceledException)
        {
            ExplorerDebugLog.Write($"Tab cleanup cancelled; no late close command was sent tab={tab.Handle}");
        }
    }
}
