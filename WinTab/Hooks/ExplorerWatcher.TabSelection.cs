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

using WindowEntry = DualKeyEntry<InternetExplorer, nint?, WindowInfo>;

/// <summary>
/// Finding an open tab for a location and switching a window to one of its tabs.
/// </summary>
public partial class ExplorerWatcher
{
    /// <summary>
    /// How long a synchronous Explorer command may take to be acknowledged. A busy Explorer often needs
    /// more than a few hundred milliseconds to activate a tab; the outcome is always verified by
    /// observing the window afterwards, never by the acknowledgement alone.
    /// </summary>
    private const int ExplorerCommandTimeoutMs = 1_000;

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
                    // Another merge's concealed or closing source is on its way out; a tab there would be lost
                    // with it, and two such sources could otherwise reuse each other and both close.
                    if (IsMergeSourceWindow(topLevelWindow))
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
                    !IsCurrentWindow(foundWindow, info) || !IsCurrentTab(info, matchedTabHandle) ||
                    IsMergeSourceWindow(info.Identity.Handle))
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
    private int _tabSelectionsInProgress;
    private int _ignoreNativeFocusThrough = Environment.TickCount;

    public async Task<bool> SelectTabByHandle(nint windowHandle, nint tabHandle, int timeoutMs = 2_500, bool bringToFront = true)
    {
        if (windowHandle == 0 || tabHandle == 0)
            return false;
        Interlocked.Increment(ref _tabSelectionsInProgress);
        try
        {
            var parentIdentity = WindowIdentity.Capture(windowHandle);
            var tabIdentity = WindowIdentity.Capture(tabHandle);
            EnsureWindowIdentity(parentIdentity);
            EnsureWindowIdentity(tabIdentity);
            ExplorerDebugLog.Write($"SelectTab hwnd={windowHandle} tab={tabHandle} foreground={WinApi.GetForegroundWindow()} bringToFront={bringToFront}");
            if (bringToFront)
            {
                var wasMinimized = WinApi.IsIconic(windowHandle);
                Helper.RestoreWindowToForeground(windowHandle);
                if (wasMinimized)
                {
                    // Restoring a minimized frame is asynchronous. Do not race its old-view focus
                    // restoration with tab-switch commands while it is still coming to the foreground.
                    var foreground = await Helper.DoUntilConditionAsync(ExplorerNavigationAccess.ForegroundFrame,
                        handle => handle == windowHandle, 500, 10, CurrentCancellation);
                    EnsureWindowIdentity(parentIdentity);
                    EnsureWindowIdentity(tabIdentity);
                    if (foreground != windowHandle)
                        return false;
                }
            }
            else if (WinApi.IsIconic(windowHandle))
                return false;
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
        finally
        {
            // Switching by index can focus intermediate tabs. Delayed WinEvents from those commands are
            // not new native file-location requests, even if delivered after this method has returned.
            Volatile.Write(ref _ignoreNativeFocusThrough, Environment.TickCount);
            Interlocked.Decrement(ref _tabSelectionsInProgress);
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
        var startedAt = Environment.TickCount64;
        if (!WinApi.TrySendMessage(windowHandle, WinApi.WM_COMMAND, 0xA221, index + 1, ExplorerCommandTimeoutMs))
        {
            var error = Marshal.GetLastWin32Error();
            ExplorerDebugLog.Write($"Tab switch command not acknowledged hwnd={windowHandle} index={index} error={error} elapsed={Environment.TickCount64 - startedAt}");
        }
    }
}
