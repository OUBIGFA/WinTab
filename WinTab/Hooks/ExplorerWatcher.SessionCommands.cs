using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Shell32;
using SHDocVw;
using WinTab.Helpers;
using WinTab.Models;

namespace WinTab.Hooks;

using WindowEntry = DualKeyEntry<InternetExplorer, nint?, WindowInfo>;

public enum SessionCommandResult
{
    Completed,
    /// <summary>Restored, but saved locations that no longer exist or are not local were skipped.</summary>
    CompletedWithSkips,
    NothingSaved,
    NothingAvailable,
    NoClosedTab,
    Busy,
    Failed,
    NotReady
}

/// <summary>
/// Restores the user asks for from the tray or a shortcut: the last tab group in a window of its own, and the
/// most recently closed tab. Both run on the shell worker, one at a time, within a bounded budget.
/// </summary>
public partial class ExplorerWatcher
{
    private const int RequestedWindowWaitMs = 8_000;
    private const int RequestedCommandBudgetMs = 60_000;
    private const int ReopenTabBudgetMs = 15_000;
    private int _independentOpens;
    private int _sessionCommandsInProgress;

    /// <summary>The user holds Ctrl + Shift, or a bounded request is opening a window that must stay on its own.</summary>
    private bool IsIndependentOpenRequested() =>
        Helper.IsCtrlShiftDown() || Volatile.Read(ref _independentOpens) != 0;

    public bool HasClosedTabs => _closedTabs.Count > 0;

    public Task<SessionCommandResult> RestoreLastSessionAsync() =>
        RunSessionCommandAsync("restore-group", RestoreLastSessionCoreAsync);

    public Task<SessionCommandResult> ReopenClosedTabAsync() =>
        RunSessionCommandAsync("reopen-tab", ReopenClosedTabCoreAsync);

    private async Task<SessionCommandResult> RunSessionCommandAsync(string name,
        Func<CancellationToken, Task<SessionCommandResult>> command)
    {
        if (_disposed || !_captureSessions || !IsShellReady)
            return SessionCommandResult.NotReady;
        if (Interlocked.CompareExchange(ref _sessionCommandsInProgress, 1, 0) != 0)
            return SessionCommandResult.Busy;
        var generation = _shellGeneration;
        var lifetime = _shellLifetime.Token;
        try
        {
            var result = await Task.Factory.StartNew(async () =>
            {
                if (_disposed || generation != _shellGeneration || lifetime.IsCancellationRequested)
                    return SessionCommandResult.NotReady;
                using var budget = CancellationTokenSource.CreateLinkedTokenSource(lifetime);
                budget.CancelAfter(RequestedCommandBudgetMs);
                return await command(budget.Token);
            }, CancellationToken.None, TaskCreationOptions.DenyChildAttach, _staTaskScheduler).Unwrap();
            ExplorerDebugLog.Write($"Session command {name} result={result}");
            return result;
        }
        catch (OperationCanceledException)
        {
            ExplorerDebugLog.Write($"Session command {name} cancelled");
            return SessionCommandResult.Failed;
        }
        catch (Exception exception) when (IsDisconnectedShell(exception))
        {
            if (!_disposed && generation == _shellGeneration)
            {
                RetireShellConnection("session-command-disconnected");
                ReportStatus($"Explorer connection disconnected ({exception.HResult:X8}); reconnecting.");
            }
            return SessionCommandResult.Failed;
        }
        catch (Exception exception)
        {
            ReportStatus($"Session command {name} failed: {exception.GetType().Name}: {exception.Message}");
            return SessionCommandResult.Failed;
        }
        finally
        {
            Volatile.Write(ref _sessionCommandsInProgress, 0);
        }
    }

    /// <summary>
    /// Opens the last group as a window of its own, beside any windows already open. Only locations the
    /// automatic restore would accept are opened; the group stays saved, so it can be restored again.
    /// </summary>
    private async Task<SessionCommandResult> RestoreLastSessionCoreAsync(CancellationToken token)
    {
        await AwaitClosedSessionsAsync(token);
        if (_sessionStore?.Snapshot is not { } session || !QualifiesAsGroup(session))
            return SessionCommandResult.NothingSaved;
        if (Volatile.Read(ref _sessionRestoresInProgress) != 0)
            return SessionCommandResult.Busy;
        var available = await _sessionLocationPolicy.FindAvailableAsync(session, token);
        if (ExplorerSessionRestorePlan.CreateForNewWindow(session, available) is not { } created)
            return SessionCommandResult.NothingAvailable;
        var (plan, first) = created;
        var location = session.Locations[first];
        ExplorerDebugLog.Write($"Requested restore saved={session.Locations.Length} available={available.Count} first={first} active={plan.ActiveSavedIndex}");
        if (await OpenRequestedWindowAsync(location, token) is not { } opened)
            return SessionCommandResult.Failed;

        var (window, info) = opened;
        var identity = info.Identity;
        var handle = identity.Handle;
        using var attempt = new SessionRestoreAttempt(token, ExplorerNavigationAccess.ForegroundFrame() == handle, info.TabIdentity);
        // The window is not the last group, or even captured, until every tab has been added.
        _restoringSessionWindows[identity] = attempt;
        _sessionTracker.Forget(identity);
        _sessionWindows.TryRemove(new KeyValuePair<nint, WindowIdentity>(handle, identity));
        try
        {
            // The one foreground activation, before anything is created in the window.
            Helper.RestoreWindowToForeground(handle);
            var foreground = await Helper.DoUntilConditionAsync(ExplorerNavigationAccess.ForegroundFrame,
                current => current == handle, 1_000, 20, attempt.Token);
            if (foreground != handle)
            {
                ExplorerDebugLog.Write($"Requested restore stopped: the new window did not reach the foreground hwnd={handle}");
                return SessionCommandResult.Failed;
            }
            attempt.ObserveForeground(true);
            if (plan.Tabs.Length > 0)
            {
                attempt.ArmNavigation();
                info.Location = location;
                if (!await RestoreSessionInWindowAsync(window, info, location, plan, _sessionGeneration, attempt.Token, requested: true))
                    return SessionCommandResult.Failed;
            }
            return plan.SkippedCount > 0 ? SessionCommandResult.CompletedWithSkips : SessionCommandResult.Completed;
        }
        catch (OperationCanceledException)
        {
            ExplorerDebugLog.Write($"Requested restore cancelled by the user or its budget hwnd={handle}");
            return SessionCommandResult.Failed;
        }
        finally
        {
            _restoringSessionWindows.TryRemove(new KeyValuePair<WindowIdentity, SessionRestoreAttempt>(identity, attempt));
            _selectionWork.Request();
        }
    }

    /// <summary>
    /// Reopens the most recently closed tab: in the Explorer window in front, else in the window it was closed
    /// in, else in the window used last, else in a window of its own. With tab reuse on, a location that is
    /// already open is brought forward instead.
    /// </summary>
    private async Task<SessionCommandResult> ReopenClosedTabCoreAsync(CancellationToken token)
    {
        if (!_recordClosedTabs)
            return SessionCommandResult.NoClosedTab;
        CompleteClosedSessions();
        if (!_closedTabs.TryPop(out var closed))
            return SessionCommandResult.NoClosedTab;
        ExplorerDebugLog.Write($"Reopen closed tab location={closed.Location}");
        var reopened = false;
        try
        {
            var available = await _sessionLocationPolicy.FindAvailableAsync(new ExplorerSession
            {
                Locations = [closed.Location], ActiveTabIndex = 0, OrderVerified = true
            }, token);
            if (!available.Contains(0))
                return SessionCommandResult.NothingAvailable;
            token.ThrowIfCancellationRequested();
            if (!_recordClosedTabs)
                return SessionCommandResult.NoClosedTab;
            reopened = await ReopenAsync(closed, token);
            return reopened ? SessionCommandResult.Completed : SessionCommandResult.Failed;
        }
        finally
        {
            // A failed attempt leaves the tab to try again rather than silently forgetting it.
            if (!reopened && _recordClosedTabs)
                _closedTabs.Return(closed);
        }
    }

    private async Task<bool> ReopenAsync(ClosedTab closed, CancellationToken token)
    {
        var target = ChooseReopenTarget(closed);
        if (target == 0)
        {
            if (await OpenRequestedWindowAsync(closed.Location, token) is not { } opened)
                return false;
            Helper.RestoreWindowToForeground(opened.Info.Identity.Handle);
            return true;
        }

        var previous = _currentMerge.Value;
        using var operation = new MergeOperation(WindowIdentity.Capture(target), _hookGeneration, token,
            () => !_disposed && _captureSessions && _recordClosedTabs, ReopenTabBudgetMs);
        _currentMerge.Value = operation;
        try
        {
            return await OpenTabNavigateWithSelection(new WindowRecord(closed.Location), target);
        }
        finally
        {
            _currentMerge.Value = previous;
        }
    }

    private nint ChooseReopenTarget(ClosedTab closed)
    {
        var foreground = ExplorerNavigationAccess.ForegroundFrame();
        if (IsReopenTarget(foreground))
            return foreground;
        if (closed.Frame.IsCurrent && IsReopenTarget(closed.Frame.Handle))
            return closed.Frame.Handle;
        if (_sessionTracker.MostRecent(_ => true) is { } recent && recent.Identity.IsCurrent && IsReopenTarget(recent.Identity.Handle))
            return recent.Identity.Handle;
        return _getExplorerWindows().FirstOrDefault(IsReopenTarget);
    }

    private bool IsReopenTarget(nint handle) =>
        handle != 0 && ExplorerWindowDiscovery.IsShownExplorerWindow(handle) && !IsMergeSourceWindow(handle) &&
        HasHookedShellWindowForTopLevel(handle) && !_restoringSessionWindows.Keys.Any(identity => identity.Handle == handle);

    /// <summary>
    /// Asks Explorer for a new window at the location and returns its tab once WinTab tracks it there. Until
    /// then new windows are treated like Ctrl + Shift windows: never merged and never restored into.
    /// </summary>
    private async Task<(InternetExplorer Window, WindowInfo Info)?> OpenRequestedWindowAsync(string location, CancellationToken token)
    {
        var known = _getExplorerWindows().Where(ExplorerWindowDiscovery.IsShownExplorerWindow).ToHashSet();
        Interlocked.Increment(ref _independentOpens);
        try
        {
            Helper.BypassWinForegroundRestrictions();
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
            }, cancellationToken: token);

            // A newly shown frame is not sufficient ownership evidence: another user launch can race
            // this request. Only an idle, sole tab at the requested location may become the restore target.
            var opened = await Helper.DoUntilNotDefaultAsync(() => _getExplorerWindows()
                .Where(handle => !known.Contains(handle) && ExplorerWindowDiscovery.IsShownExplorerWindow(handle))
                .Select(FindRequestedWindowTab)
                .FirstOrDefault(candidate => candidate is { } tab && IsSessionWindowIdle(tab.Window) &&
                    ExplorerSessionRestorePlan.SameLocation(TryGetLocation(tab.Window), location) &&
                    ExplorerWindowDiscovery.GetAllExplorerTabs(tab.Info.Identity.Handle).Take(2).Count() == 1),
                RequestedWindowWaitMs, 25, token);
            if (opened is not { } result)
            {
                ReportStatus("Explorer did not confirm a new window at the requested restore location.");
                return null;
            }
            token.ThrowIfCancellationRequested();
            var handle = result.Info.Identity.Handle;
            ExcludeWindowFromSessionRestore(handle);
            PreventWindowHiding(handle);
            await RestoreMergeSourceWindowAsync(handle);
            token.ThrowIfCancellationRequested();
            if (!IsCurrentWindow(result.Window, result.Info) || !result.Info.TabIdentity.IsCurrent)
                return null;
            return result;
        }
        finally
        {
            Interlocked.Decrement(ref _independentOpens);
        }
    }

    private (InternetExplorer Window, WindowInfo Info)? FindRequestedWindowTab(nint handle)
    {
        lock (_windowEntryDictLock)
        {
            foreach (var entry in (IEnumerable<WindowEntry>)_windowEntryDict)
            {
                var info = entry.Value;
                if (info.Identity.Handle == handle && info.EventsHooked && entry.OptionalKey is { } tab &&
                    IsCurrentTab(info, tab) && IsCurrentWindow(entry.PrimaryKey, info))
                    return (entry.PrimaryKey, info);
            }
        }
        return null;
    }
}
