using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.WinAPI;
using Fixture = ExplorerTabLifetimeTests.Fixture;

/// <summary>
/// Windows 11 25H2 preloads a hidden File Explorer frame seconds after a folder opens and reuses that
/// frame for the next folder the user opens. These tests reproduce what happened on the desktop: the
/// merge budget ran while the frame was still hidden, the frame was then restored and protected as if it
/// were a user window, and slow command acknowledgements from a busy Explorer were treated as failures.
/// </summary>
internal static class ExplorerPreloadedFrameTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("the merge budget of a hidden Explorer frame starts when Explorer shows it", BudgetStartsWhenShown);
        yield return ("a frame Explorer keeps hidden is neither restored nor reported as a timed-out merge", HiddenFrameIsNotExpired);
        yield return ("a preloaded frame shown after the old budget stays an unprotected merge source", ShownPreloadedFrameStaysMergeable);
        yield return ("a hidden Explorer frame is never chosen as a merge target", HiddenFrameIsNotAMergeTarget);
        yield return ("hidden Explorer frames are concealed when merging starts", HiddenFramesAreConcealedOnStart);
        yield return ("frame maintenance conceals preloads created after the last event pulse", LatePreloadIsConcealed);
        yield return ("frame maintenance repairs opacity resets throughout preload lifetime", PreloadOpacityIsMaintained);
        yield return ("hidden catalog frames wait for their first show before registration", HiddenCatalogFrameWaitsForShow);
        yield return ("recycled frames do not inherit an old registration hiding exemption", RecycledFrameDoesNotInheritRegistration);
        yield return ("recycled frames do not inherit the old main window hiding exemption", RecycledFrameDoesNotInheritMainWindow);
        yield return ("native frame maintenance leaves current user windows and stopped hooks alone", MaintenancePreservesUserWindows);
        yield return ("frame maintenance keeps pending closes concealed past the merge deadline", PendingCloseMaintenanceOutlivesDeadline);
        yield return ("continuous preload maintenance cannot postpone expired source recovery", MaintenanceCannotStarveRecovery);
        yield return ("a hidden Explorer frame does not cause periodic COM scans", HiddenFrameDoesNotRescanCatalog);
        yield return ("a source that closes after a slow close acknowledgement counts as merged", SlowCloseAcknowledgementStillCounts);
        yield return ("a source that stays open after the close request is restored", UnclosedSourceIsRestored);
        yield return ("a busy merge source cannot execute a close after recovery", BusySourceCannotCloseAfterRecovery);
        yield return ("a source still handling its close is neither shown nor reported as restored", PendingCloseKeepsSourceConcealed);
        yield return ("a source that answers after a pending close is restored without a second close", PendingCloseThatIsRefusedIsRestoredLater);
        yield return ("stopping with a pending close restores the window without resending the close", StoppingWithPendingCloseRestoresWithoutResending);
        yield return ("a slow tab-switch acknowledgement is not a failed switch", SlowTabSwitchStillSucceeds);
    }

    private static Task BudgetStartsWhenShown() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var handle = fixture.Window.Handle;
        fixture.EnableMerging();
        Check.That(!WinApi.IsWindowVisible(handle), "The isolated frame must start hidden like a preloaded Explorer frame.");

        fixture.Invoke("HideMergeSourceWindow", handle);
        Check.That(ExplorerWindowVisibility.Contains(handle), "A preloaded frame is concealed early so it cannot flash when Explorer shows it.");
        await Task.Delay(700);
        Check.Equal(Fixture.MergeTimeoutMs, fixture.RemainingMergeTime(handle),
            "The merge budget must not run while Explorer still keeps the frame hidden.");

        fixture.Window.Show();
        fixture.Invoke("HideMergeSourceWindow", handle);
        await Task.Delay(700);
        var remaining = fixture.RemainingMergeTime(handle);
        Check.That(remaining < Fixture.MergeTimeoutMs - 500 && remaining > Fixture.MergeTimeoutMs - 2_000,
            $"The merge budget must start counting once Explorer shows the frame; remaining={remaining}.");
    }, tabCount: 1);

    private static Task HiddenFrameIsNotExpired() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var handle = fixture.Window.Handle;
        fixture.EnableMerging();
        fixture.Invoke("HideMergeSourceWindow", handle);
        await Task.Delay(Fixture.MergeTimeoutMs + 400);

        fixture.RunMergeSafetyTimer();

        Check.That(ExplorerWindowVisibility.Contains(handle), "A frame Explorer never showed must stay concealed for its eventual use.");
        Check.Equal(1, fixture.MergeSourceCount, "The hidden frame must remain a pending merge source.");
        Check.That(!(bool)fixture.Invoke("IsWindowProtected", handle)!, "A hidden frame must not be protected as an already-handled user window.");
        Check.Equal(0, fixture.Statuses.Count, "A hidden frame must not raise a merge warning: " + string.Join(" | ", fixture.Statuses));
    }, tabCount: 1);

    private static Task ShownPreloadedFrameStaysMergeable() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        var handle = fixture.Window.Handle;
        fixture.EnableMerging();
        fixture.Invoke("HideMergeSourceWindow", handle);
        await Task.Delay(Fixture.MergeTimeoutMs + 400);
        fixture.RunMergeSafetyTimer();

        // Explorer now shows the preloaded frame for the folder the user double-clicked.
        fixture.Window.Show();
        fixture.Invoke("HideMergeSourceWindow", handle);

        Check.That(!(bool)fixture.Invoke("IsWindowProtected", handle)!,
            "A preloaded frame shown seconds later must still be merged, not released as an already-handled window.");
        Check.That(fixture.RemainingMergeTime(handle) > Fixture.MergeTimeoutMs / 2,
            "A shown preloaded frame must get a full merge budget instead of inheriting an expired one.");
        Check.That(ExplorerWindowVisibility.Contains(handle), "The shown frame must remain concealed until the merge completes.");
    }, tabCount: 1);

    private static Task HiddenFrameIsNotAMergeTarget() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var hidden = new RemoteExplorerFrame(visible: false, explorerClass: true);
        using var shown = new RemoteExplorerFrame(visible: true, explorerClass: true);

        fixture.SetExplorerWindows(hidden.Handle);
        Check.Equal((nint)0, (nint)fixture.Invoke("GetMainWindowHWnd", (nint)0, null)!,
            "A frame Explorer has not shown yet must never receive merged tabs.");

        fixture.SetExplorerWindows(hidden.Handle, shown.Handle);
        Check.Equal(shown.Handle, (nint)fixture.Invoke("GetMainWindowHWnd", (nint)0, null)!,
            "A shown Explorer window must still be an acceptable fallback target.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task HiddenFramesAreConcealedOnStart() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var hidden = new RemoteExplorerFrame(visible: false, explorerClass: true);
        using var shown = new RemoteExplorerFrame(visible: true, explorerClass: true);
        fixture.SetExplorerWindows(hidden.Handle, shown.Handle);
        fixture.EnableMerging();

        fixture.Invoke("ConcealPreloadedExplorerFrames");

        Check.That(ExplorerWindowVisibility.Contains(hidden.Handle), "A preloaded frame found at start must be concealed before Explorer can show it.");
        Check.Equal(Fixture.MergeTimeoutMs, fixture.RemainingMergeTime(hidden.Handle), "A frame concealed at start must not spend its merge budget while hidden.");
        Check.That(!ExplorerWindowVisibility.Contains(shown.Handle), "A window that is already on screen must be left alone at start.");
        fixture.Invoke("RecoverHiddenExplorerWindows", "test-start-cleanup");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task LatePreloadIsConcealed() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        fixture.EnableMerging();
        fixture.SetExplorerWindows();
        fixture.MaintainFrameConcealment();
        await Task.Delay(1_600); // No running event pulse when Explorer creates this replacement preload.
        using var frame = new RemoteExplorerFrame(visible: false, explorerClass: true);
        fixture.SetExplorerWindows(frame.Handle);
        fixture.MaintainFrameConcealment();
        CheckConcealed(frame.Handle, "A missed create event must not leave a late preload unprotected until SHOW.");
        Check.Equal(Fixture.MergeTimeoutMs, fixture.RemainingMergeTime(frame.Handle), "Maintenance must not spend the hidden frame's merge budget.");
        fixture.Invoke("RecoverHiddenExplorerWindows", "test-late-preload-cleanup");
    }, tabCount: 1);

    private static Task PreloadOpacityIsMaintained() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: false, explorerClass: true);
        fixture.EnableMerging();
        fixture.SetExplorerWindows(frame.Handle);
        fixture.MaintainFrameConcealment();
        for (var cycle = 0; cycle < 100; cycle++)
        {
            // Explorer can rewrite the style or alpha without another create/show notification.
            if (cycle % 2 == 0)
                ExplorerWindowVisibility.UpdateLayeredStyle(frame.Handle, remove: true);
            else
                WinApi.SetLayeredWindowAttributes(frame.Handle, 0, 255, WinApi.LWA_ALPHA);
            fixture.MaintainFrameConcealment();
            CheckConcealed(frame.Handle, $"Preload opacity must remain maintained on cycle {cycle}.");
            Check.Equal(Fixture.MergeTimeoutMs, fixture.RemainingMergeTime(frame.Handle), "Idle maintenance must not start a merge.");
        }
        fixture.Invoke("RecoverHiddenExplorerWindows", "test-preload-opacity-cleanup");
        Check.That((WinApi.GetWindowLong(frame.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0,
            "Maintenance must retain the original recovery snapshot rather than the intermediate transparent style.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task HiddenCatalogFrameWaitsForShow() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: false, explorerClass: true);
        fixture.EnableMerging();
        fixture.SetExplorerWindows(frame.Handle);
        fixture.MaintainFrameConcealment();
        var locationReads = 0;
        var browser = fixture.AddBrowser(out _, track: false, handle: frame.Handle, readLocation: () => locationReads++);
        fixture.SetCatalog(browser);
        await (Task)fixture.Invoke("ProcessRegisteredShellWindowsAsync")!;
        await Task.Delay(Fixture.MergeTimeoutMs + 100);
        fixture.RunMergeSafetyTimer();
        Check.Equal(0, fixture.Count, "A hidden catalog frame is not an independent window or an active merge.");
        Check.Equal(0, locationReads, "Preloading must not resolve or open a folder before the user opens it.");
        CheckConcealed(frame.Handle, "Catalog publication must not undo the preload's concealment.");
        Check.That(!(bool)fixture.Invoke("IsWindowProtected", frame.Handle)!, "The later folder launch must not inherit a protection grace period.");

        WinApi.ShowWindow(frame.Handle, WinApi.SW_SHOWNOACTIVATE);
        fixture.MaintainFrameConcealment();
        fixture.Invoke("AdoptNewShellWindows");
        Check.Equal(1, fixture.Count, "The first show makes the catalog frame eligible for registration.");
        CheckConcealed(frame.Handle, "The first show must remain concealed while registration catches up.");
        fixture.Invoke("RecoverHiddenExplorerWindows", "test-hidden-catalog-cleanup");
    }, tabCount: 1);

    private static Task RecycledFrameDoesNotInheritRegistration() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: false, explorerClass: true);
        fixture.EnableMerging();
        var oldBrowser = fixture.AddBrowser(out var oldInfo, frame.Tab, handle: frame.Handle);
        fixture.Invoke("HookWindowEvents", oldBrowser, oldInfo);
        Check.That(!(bool)fixture.Invoke("TryHideIncomingExplorerWindow", frame.Handle)!, "The live original registration must be protected.");
        oldInfo.Identity.Release(); // Same HWND/thread/process but a different native lifetime token.
        Check.That(WindowIdentity.Capture(frame.Handle).IsCurrent, "The replacement identity must be current.");
        Check.That((bool)fixture.Invoke("TryHideIncomingExplorerWindow", frame.Handle)!,
            "A stale registration must not exempt a replacement from early concealment.");
        CheckConcealed(frame.Handle, "The replacement must be hidden before catalog cleanup.");
        fixture.Invoke("RecoverHiddenExplorerWindows", "test-recycled-registration-cleanup");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task RecycledFrameDoesNotInheritMainWindow() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: false, explorerClass: true);
        fixture.EnableMerging();
        fixture.CacheMainWindow(frame.Handle);
        Check.That(!(bool)fixture.Invoke("TryHideIncomingExplorerWindow", frame.Handle)!, "The live main window must be protected.");
        WindowIdentity.Capture(frame.Handle).Release();
        Check.That((bool)fixture.Invoke("TryHideIncomingExplorerWindow", frame.Handle)!,
            "A main window's cached HWND must not protect a different window lifetime.");
        CheckConcealed(frame.Handle, "A recycled main HWND must be concealed before registration.");
        fixture.Invoke("RecoverHiddenExplorerWindows", "test-recycled-main-cleanup");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task MaintenancePreservesUserWindows() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var shown = new RemoteExplorerFrame(visible: true, explorerClass: true);
        using var registered = new RemoteExplorerFrame(visible: false, explorerClass: true);
        using var preload = new RemoteExplorerFrame(visible: false, explorerClass: true);
        fixture.EnableMerging();
        fixture.MarkHooked(registered.Handle);
        fixture.SetExplorerWindows(shown.Handle, registered.Handle);
        fixture.MaintainFrameConcealment();
        Check.That(!ExplorerWindowVisibility.Contains(shown.Handle), "Maintenance must not claim existing shown windows.");
        Check.That(!ExplorerWindowVisibility.Contains(registered.Handle), "A live registered window remains protected while natively hidden.");
        fixture.DisableMerging();
        fixture.SetExplorerWindows(preload.Handle);
        fixture.MaintainFrameConcealment();
        Check.That(!ExplorerWindowVisibility.Contains(preload.Handle), "A stopped hook must not conceal new preloads.");
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task PendingCloseMaintenanceOutlivesDeadline() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        fixture.EnableMerging();
        fixture.Invoke("HideMergeSourceWindow", frame.Handle);
        Check.That((bool)fixture.Invoke("MarkClosePending", frame.Handle)!, "The source must be marked as handling its close.");
        await Task.Delay(Fixture.MergeTimeoutMs + 100);
        WinApi.SetLayeredWindowAttributes(frame.Handle, 0, 255, WinApi.LWA_ALPHA);
        fixture.MaintainFrameConcealment();
        CheckConcealed(frame.Handle, "A pending close must remain concealed even after the original merge deadline.");
        fixture.Invoke("RecoverHiddenExplorerWindows", "test-pending-deadline-cleanup");
    }, tabCount: 1);

    private static Task MaintenanceCannotStarveRecovery() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var preload = new RemoteExplorerFrame(visible: false, explorerClass: true);
        using var shown = new RemoteExplorerFrame(visible: true, explorerClass: true);
        using var safetyTimer = fixture.StartMergeSafetyTimer();
        fixture.EnableMerging();
        fixture.SetExplorerWindows(preload.Handle);
        fixture.Invoke("HideMergeSourceWindow", shown.Handle);
        var until = Environment.TickCount64 + Fixture.MergeTimeoutMs + 750;
        while (Environment.TickCount64 < until)
        {
            fixture.MaintainFrameConcealment();
            await Task.Delay(20);
        }
        Check.That(!ExplorerWindowVisibility.Contains(shown.Handle),
            "Frequent re-hides must not keep restarting the recovery timer before it can run.");
        CheckConcealed(preload.Handle, "Recovering an expired source must leave the unused preload concealed.");
        fixture.Invoke("RecoverHiddenExplorerWindows", "test-recovery-timer-cleanup");
    }, tabCount: 1);

    private static void CheckConcealed(nint handle, string message) => Check.That(
        ExplorerWindowVisibility.Contains(handle) &&
        WinApi.GetLayeredWindowAttributes(handle, out _, out var alpha, out var flags) &&
        (flags & WinApi.LWA_ALPHA) != 0 && alpha == 0, message);

    private static Task HiddenFrameDoesNotRescanCatalog() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        Check.That(!WinApi.IsWindowVisible(fixture.Window.Handle), "The isolated frame must start hidden.");
        Check.That(!await fixture.PollForRegistrationAsync(fixture.Window.Handle),
            "A hidden preloaded frame cannot be registered yet and must not keep full catalog scans running every second.");

        fixture.Window.Show();
        Check.That(await fixture.PollForRegistrationAsync(fixture.Window.Handle),
            "Once Explorer shows the frame its untracked tab must be discovered.");
    }, tabCount: 1);

    private static Task SlowCloseAcknowledgementStillCounts() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true) { CloseDelayMs = 450 };
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);

        var closed = await fixture.CloseMergedSourceAsync(browser, info);

        Check.That(closed, "A source that closes while Explorer is still acknowledging the close request must count as merged.");
        Check.That(!frame.IsAlive, "The source window must actually be gone.");
        Check.That(!fixture.Statuses.Any(status => status.Contains("did not close", StringComparison.OrdinalIgnoreCase)),
            "A slow close must not be reported as a failed close: " + string.Join(" | ", fixture.Statuses));
    }, tabCount: 1);

    private static Task UnclosedSourceIsRestored() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true) { IgnoreClose = true };
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);

        var closed = await fixture.CloseMergedSourceAsync(browser, info);

        Check.That(!closed, "A source that refuses to close must not be reported as merged.");
        Check.That(frame.IsAlive, "The refusing source window must remain open.");
        Check.That(!ExplorerWindowVisibility.Contains(frame.Handle), "A source that stays open must leave the concealed list.");
        Check.That((WinApi.GetWindowLong(frame.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0,
            "A source that stays open must be made visible again for the user.");
    }, tabCount: 1);

    private static async Task SlowTabSwitchStillSucceeds()
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            using var lifetime = new CancellationTokenSource();
            var watcher = ExplorerTabActivationTests.CreateSelectionWatcher(lifetime);
            using var frame = new RemoteExplorerFrame(visible: true, tabCount: 2) { SwitchDelayMs = 450 };
            frame.SetActive(1);
            Check.That(frame.ActiveTab != frame.Tab, "The test must start on another tab.");

            var selected = await watcher.SelectTabByHandle(frame.Handle, frame.Tab, timeoutMs: 2_500);

            Check.That(selected, $"A slow switch must succeed; target={frame.Tab}, active={frame.ActiveTab}, commands={string.Join(',', frame.SwitchCommands)}, trace={frame.Trace}");
            Check.Equal(frame.Tab, frame.ActiveTab, "The requested tab must be active after the slow switch.");
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }

    /// <summary>
    /// Explorer is busy (not pumping messages) when the close is requested. The timed close request is
    /// discarded by Windows, so nothing may close the window later; it is restored once Explorer answers.
    /// </summary>
    private static Task BusySourceCannotCloseAfterRecovery() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true);
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);
        fixture.EnableMerging();
        fixture.Invoke("HideMergeSourceWindow", frame.Handle);
        await frame.BlockMessagesAsync(2_200);
        try
        {
            var closed = await fixture.CloseMergedSourceAsync(browser, info, recover: false);

            Check.That(!closed, "A source that cannot process the close within the deadline must not count as closed.");
            Check.That(!ReportedRestored(fixture), "A window that has not answered since the close request must not be reported as restored: " + Statuses(fixture));
            fixture.RunMergeSafetyTimer();
            Check.That(ExplorerWindowVisibility.Contains(frame.Handle), "The source stays concealed until Explorer answers again.");

            await Task.Delay(1_500);
            fixture.RunMergeSafetyTimer();
            await Task.Delay(150);

            Check.That(frame.IsAlive, "Recovering the source must not leave a close request that destroys it afterwards.");
            Check.Equal(0, frame.CloseCommandCount, "The busy frame must never receive a stale close after it resumes.");
            Check.That(!ExplorerWindowVisibility.Contains(frame.Handle), "The failed merge source must be restored once Explorer answers.");
            Check.That(ReportedRestored(fixture), "The recovery must be reported once it has happened: " + Statuses(fixture));
        }
        finally
        {
            fixture.Invoke("RecoverHiddenExplorerWindows", "test-busy-cleanup");
        }
    }, tabCount: 1);

    /// <summary>Explorer picks the close up at once but takes longer than the observation window to destroy the frame.</summary>
    private static Task PendingCloseKeepsSourceConcealed() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true) { CloseDelayMs = 1_500 };
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);
        try
        {
            var closed = await fixture.CloseMergedSourceAsync(browser, info, recover: false);

            Check.That(!closed, "An unfinished close must not be reported as merged; the window stays tracked until Explorer answers.");
            Check.That(frame.IsAlive && ExplorerWindowVisibility.Contains(frame.Handle), "A source still handling its close stays concealed.");
            fixture.RunMergeSafetyTimer();
            Check.That(ExplorerWindowVisibility.Contains(frame.Handle), "The safety timer must not show a window that has not answered since its close request.");

            await Task.Delay(700);
            Check.That(!frame.IsAlive, "The frame must have finished closing by now.");
            Check.That(!frame.RevealedWhileClosing, "A window that is still being destroyed must never be shown again.");
            fixture.RunMergeSafetyTimer();

            Check.Equal(0, fixture.MergeSourceCount, "A source that finished closing is forgotten.");
            Check.That(!ExplorerWindowVisibility.Contains(frame.Handle), "A closed source leaves the concealed list.");
            Check.Equal(1, frame.CloseCommandCount, "A close that is still being handled must not be repeated.");
            Check.That(!ReportedRestored(fixture), "A source that closed must never be reported as restored: " + Statuses(fixture));
        }
        finally
        {
            fixture.Invoke("RecoverHiddenExplorerWindows", "test-pending-cleanup");
        }
    }, tabCount: 1);

    /// <summary>Explorer picks the close up, stays silent for a while, then keeps the window open.</summary>
    private static Task PendingCloseThatIsRefusedIsRestoredLater() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true) { CloseDelayMs = 1_500, IgnoreClose = true };
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);
        try
        {
            var closed = await fixture.CloseMergedSourceAsync(browser, info, recover: false);

            Check.That(!closed && ExplorerWindowVisibility.Contains(frame.Handle), "A source still handling its close stays concealed and is not reported as closed.");
            // Explorer shows the frame again while the close is pending: it is re-hidden, never closed a second time.
            fixture.Invoke("TryHideIncomingExplorerWindow", frame.Handle);
            Check.Equal(1, frame.CloseCommandCount, "A source shown again during a pending close must not receive another close request.");

            await Task.Delay(600);
            fixture.RunMergeSafetyTimer();
            await Task.Delay(100);

            Check.That(frame.IsAlive, "The window Explorer kept open must survive.");
            Check.That(!ExplorerWindowVisibility.Contains(frame.Handle), "The window is restored once Explorer answers again.");
            Check.That((WinApi.GetWindowLong(frame.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0, "The restored window must be visible again for the user.");
            Check.Equal(1, frame.CloseCommandCount, "A window that answered after a pending close must not be closed again.");
            Check.That(fixture.Statuses.Any(status => status.Contains("kept", StringComparison.OrdinalIgnoreCase)),
                "The recovery must report that Explorer kept the window: " + Statuses(fixture));
        }
        finally
        {
            fixture.Invoke("RecoverHiddenExplorerWindows", "test-refused-cleanup");
        }
    }, tabCount: 1);

    /// <summary>Stopping WinTab while a close is still being handled leaves no hidden window behind and sends no second close.</summary>
    private static Task StoppingWithPendingCloseRestoresWithoutResending() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true) { CloseDelayMs = 1_500, IgnoreClose = true };
        var browser = fixture.AddBrowser(out var info, frame.Tab, handle: frame.Handle);

        var closed = await fixture.CloseMergedSourceAsync(browser, info, recover: false);
        Check.That(!closed && ExplorerWindowVisibility.Contains(frame.Handle), "The close must still be pending when WinTab stops.");
        Check.That(!ReportedRestored(fixture), "A pending close must not be reported as restored before stopping: " + Statuses(fixture));

        fixture.Invoke("RecoverHiddenExplorerWindows", "stop-hook");
        await Task.Delay(150);

        Check.That(frame.IsAlive, "The window Explorer kept open must survive the stop.");
        Check.That(!ExplorerWindowVisibility.Contains(frame.Handle) && fixture.MergeSourceCount == 0, "Stopping must leave no concealed window behind.");
        Check.That((WinApi.GetWindowLong(frame.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0, "The window must be visible again after stopping.");
        Check.Equal(1, frame.CloseCommandCount, "Stopping must not send another close request.");
    }, tabCount: 1);

    private static bool ReportedRestored(Fixture fixture) =>
        fixture.Statuses.Any(status => status.Contains("restored", StringComparison.OrdinalIgnoreCase));

    private static string Statuses(Fixture fixture) => string.Join(" | ", fixture.Statuses);
}
