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
        yield return ("a hidden Explorer frame does not cause periodic COM scans", HiddenFrameDoesNotRescanCatalog);
        yield return ("a source that closes after a slow close acknowledgement counts as merged", SlowCloseAcknowledgementStillCounts);
        yield return ("a source that stays open after the close request is restored", UnclosedSourceIsRestored);
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

            Check.That(selected, "A tab switch that Explorer acknowledges slowly must still be reported as successful.");
            Check.Equal(frame.Tab, frame.ActiveTab, "The requested tab must be active after the slow switch.");
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }
}
