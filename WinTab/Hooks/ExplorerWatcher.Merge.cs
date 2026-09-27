using SHDocVw;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Concealing new windows while they are merged, choosing the window they merge into, and closing or
/// restoring them afterwards.
/// </summary>
public partial class ExplorerWatcher
{
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

    private bool TryHideIncomingExplorerWindow(nint hWnd)
    {
        hWnd = GetExplorerTopLevelWindow(hWnd);
        if (!_isForcingTabs || hWnd == 0) return false;
        if (_closingMergeSourceHWnds.ContainsKey(hWnd))
        {
            HideMergeSourceWindow(hWnd);
            return true;
        }

        if (IsWindowProtected(hWnd)) return false;
        if (IsIndependentOpenRequested()) return false;
        if (_hookedTopLevelUseCounts.ContainsKey(hWnd)) return false;
        if (_mainWindowHandle != 0 && hWnd == _mainWindowHandle) return false;
        if (ExplorerWindowDiscovery.GetAllExplorerTabs(hWnd).Take(2).Count() > 1) return false;
        if (ReleaseTornOffTabWindow(hWnd)) return false;

        HideMergeSourceWindow(hWnd);
        if (_tabTearOff.IsPending(Environment.TickCount64))
            _ = ReleaseTornOffTabWindowWhenReadyAsync(hWnd);
        return true;
    }
    /// <summary>
    /// The window Explorer opens for a tab the user dragged off the tab row is what the user asked for, so it
    /// is never merged back: it is shown again if it was concealed while Explorer still had it hidden, and it
    /// is protected like a window opened with Ctrl + Shift.
    /// </summary>
    private bool ReleaseTornOffTabWindow(nint hWnd)
    {
        if (!_tabTearOff.TryClaim(hWnd, Environment.TickCount64))
            return false;

        ExcludeWindowFromSessionRestore(hWnd);
        if (_mergeSourceHWnds.TryGetValue(hWnd, out var concealed))
        {
            if (RestoreConcealedWindow(concealed))
                ExplorerDebugLog.Write($"Torn-off tab window shown again hwnd={hWnd}");
        }
        else if (!IsWindowProtected(hWnd))
        {
            PreventWindowHiding(hWnd);
            ExplorerDebugLog.Write($"Torn-off tab window released hwnd={hWnd}");
        }
        return true;
    }

    /// <summary>
    /// Explorer may show the window for a torn-off tab a moment before it takes the tab out of the source
    /// window. Such a window is concealed like any other new window and shown again as soon as the drop can
    /// be confirmed, without waiting for its registration to catch up.
    /// </summary>
    private async Task ReleaseTornOffTabWindowWhenReadyAsync(nint hWnd)
    {
        if (!_tornOffTabReleaseWatches.TryAdd(hWnd, true))
            return;
        try
        {
            while (_isForcingTabs && !_disposed && _tabTearOff.IsPending(Environment.TickCount64) &&
                   _mergeSourceHWnds.TryGetValue(hWnd, out var concealed) && concealed.Identity.IsCurrent)
            {
                if (ReleaseTornOffTabWindow(hWnd))
                    return;
                await Task.Delay(25);
            }
        }
        catch (Exception exception)
        {
            ExplorerDebugLog.Write($"Torn-off tab window release failed hwnd={hWnd} error={exception.GetType().Name}");
        }
        finally
        {
            _tornOffTabReleaseWatches.TryRemove(hWnd, out _);
        }
    }

    /// <summary>
    /// While a drop may still produce its window, a new window cannot be told apart from the one Explorer
    /// opens for the dragged tab until Explorer has moved the tab. Its merge therefore waits until the drop
    /// is confirmed for this window, or until the drop can no longer produce a window.
    /// </summary>
    private async Task<bool> WaitForTornOffTabWindowAsync(nint hWnd)
    {
        while (!ReleaseTornOffTabWindow(hWnd))
        {
            if (!_tabTearOff.IsPending(Environment.TickCount64))
                return false;
            await Task.Delay(25, CurrentCancellation);
            EnsureCurrentMerge();
        }

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
        if (IsIndependentOpenRequested()) return;
        if (_mainWindowHandle != 0 && hWnd == _mainWindowHandle) return;
        if (ExplorerWindowDiscovery.GetAllExplorerTabs(hWnd).Take(2).Count() > 1) return;

        var targetWindow = GetMainWindowHWnd(hWnd);
        if (targetWindow == 0 || hWnd == targetWindow) return;
        if (ReleaseTornOffTabWindow(hWnd)) return;

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
        if (IsClosePending(concealed))
        {
            ExplorerDebugLog.Write($"Source restore deferred; Explorer has not answered since the close request hwnd={handle}");
            return;
        }
        var restored = await Helper.DoUntilConditionAsync(() => RestoreConcealedWindow(concealed),
            result => result, 1_000, 50);
        if (!restored)
            ReportStatus("Explorer rejected window recovery; recovery will be retried.");
    }

    /// <summary>Whether the window is a concealed merge source or one whose close has been requested.</summary>
    private bool IsMergeSourceWindow(nint handle) =>
        handle != 0 &&
        ((_mergeSourceHWnds.TryGetValue(handle, out var concealed) && concealed.Identity.IsCurrent) ||
         (_closingMergeSourceHWnds.TryGetValue(handle, out var closing) && closing.Identity.IsCurrent));

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
        var closePending = false;
        try
        {
            for (var attempt = 0; attempt < 2; attempt++)
            {
                EnsureCurrentMerge();
                if (!RequestCloseMergedSourceWindow(handle))
                    break;
                bool closed;
                try
                {
                    closed = await Helper.DoUntilConditionAsync(() => !identity.IsCurrent,
                        isClosed => isClosed, attempt == 0 ? 700 : 300, 40, CurrentCancellation);
                }
                catch (OperationCanceledException)
                {
                    closePending = identity.IsCurrent && !IsWindowAnswering(handle) && MarkClosePending(handle);
                    throw;
                }
                if (closed)
                    return true;
                // A window that answers again has finished with the request and kept itself open. A silent one
                // may still be destroying itself, so it stays concealed until it answers or disappears.
                if (!IsWindowAnswering(handle))
                {
                    closePending = MarkClosePending(handle);
                    break;
                }
            }
        }
        finally
        {
            _closingMergeSourceHWnds.TryRemove(new KeyValuePair<nint, MergeOperation>(handle, completion));
            try
            {
                if (!closePending && identity.IsCurrent)
                    await RestoreMergeSourceWindowAsync(handle);
            }
            finally
            {
                _currentMerge.Value = operation;
            }
        }

        // Recovery may pump window messages; verify the final state before reporting a failed close.
        if (!identity.IsCurrent)
            return true;
        ReportStatus(closePending
            ? "Explorer has not finished closing the source window; it stays concealed until Explorer answers."
            : "Explorer did not close the source window; it has been restored.");
        return false;
    }

    /// <summary>Asks Explorer to close the concealed source; true when a request was issued and must be observed.</summary>
    private bool RequestCloseMergedSourceWindow(nint handle)
    {
        if (!_isForcingTabs || !_closingMergeSourceHWnds.TryGetValue(handle, out var operation) || !operation.IsCurrent ||
            ExplorerWindowDiscovery.GetAllExplorerTabs(handle).Take(2).Count() > 1)
            return false;
        // Windows discards a timed send that a busy window has not picked up, so nothing can execute once the
        // window is restored; a posted WM_CLOSE would. Explorer destroys the frame while handling the request,
        // which makes the send report a failure for a successful close, so the caller observes the window.
        if (!WinApi.TrySendMessage(handle, WinApi.WM_CLOSE, 0, 0) && operation.Identity.IsCurrent)
            ExplorerDebugLog.Write($"Source close not acknowledged hwnd={handle}; observing completion");
        return true;
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
}
