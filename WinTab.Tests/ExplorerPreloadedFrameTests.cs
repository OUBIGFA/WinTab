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
        yield return ("a frame Explorer keeps hidden is neither restored nor reported as a timed-out merge", HiddenFrameIsNotExpired);
        yield return ("frame maintenance conceals preloads created after the last event pulse", LatePreloadIsConcealed);
        yield return ("frame maintenance repairs opacity resets throughout preload lifetime", PreloadOpacityIsMaintained);
        yield return ("recycled frames do not inherit an old registration hiding exemption", RecycledFrameDoesNotInheritRegistration);
        yield return ("recycled frames do not inherit the old main window hiding exemption", RecycledFrameDoesNotInheritMainWindow);
    }

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
                BackgroundWindowVisibility.UpdateLayeredStyle(frame.Handle, remove: true);
            else
                TestWindowOpacity.Instance.TryWrite(frame.Handle, 0, 255, WinApi.LWA_ALPHA);
            fixture.MaintainFrameConcealment();
            CheckConcealed(frame.Handle, $"Preload opacity must remain maintained on cycle {cycle}.");
            Check.Equal(Fixture.MergeTimeoutMs, fixture.RemainingMergeTime(frame.Handle), "Idle maintenance must not start a merge.");
        }
        fixture.Invoke("RecoverHiddenExplorerWindows", "test-preload-opacity-cleanup");
        Check.That((TestWindowOpacity.Instance.ReadStyle(frame.Handle) & WinApi.WS_EX_LAYERED) == 0,
            "Maintenance must retain the original recovery snapshot rather than the intermediate transparent style.");
        return Task.CompletedTask;
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

    private static void CheckConcealed(nint handle, string message) => Check.That(
        ExplorerWindowVisibility.Contains(handle) &&
        TestWindowOpacity.Instance.TryRead(handle, out _, out var alpha, out var flags) &&
        (flags & WinApi.LWA_ALPHA) != 0 && alpha == 0, message);

    private static bool ReportedRestored(Fixture fixture) =>
        fixture.Statuses.Any(status => status.Contains("restored", StringComparison.OrdinalIgnoreCase));

    private static string Statuses(Fixture fixture) => string.Join(" | ", fixture.Statuses);
}
