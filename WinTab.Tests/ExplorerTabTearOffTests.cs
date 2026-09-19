using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.WinAPI;
using Fixture = ExplorerTabLifetimeTests.Fixture;

/// <summary>
/// Dragging a tab off the tab row makes Explorer open a new window for it. That window is what the user
/// asked for, so WinTab leaves Explorer's native drag alone and never merges the result back. These
/// tests describe the drags Explorer turns into a window and the ones it does not, and check that the
/// watcher releases exactly the window that came out of the drag.
/// </summary>
internal static class ExplorerTabTearOffTests
{
    private const nint Source = 0x10;
    private const nint NewWindow = 0x20;
    private const nint Other = 0x30;
    private static readonly Rectangle SourceBounds = Rectangle.FromLTRB(100, 100, 1100, 750);
    private static readonly Rectangle OtherBounds = Rectangle.FromLTRB(1200, 100, 2200, 750);
    private static readonly Point OnTab = new(200, 126);
    private static readonly Point OnSecondTab = new(450, 126);
    private static readonly Point BesideTabs = new(700, 126);
    private static readonly Point BelowTabs = new(200, 400);
    private static readonly Point OutsideWindows = new(1300, 900);

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("a tab dragged off the tab row releases the window Explorer opens for it", TornOffTabReleasesNewWindow);
        yield return ("dropping a tab back on its own tab row is a reorder, not a tear-off", DropOnOwnRowIsNotATearOff);
        yield return ("dropping a tab on another window's tab row moves it, not tears it off", DropOnAnotherRowIsNotATearOff);
        yield return ("a press beside the tab titles is a window drag, not a tab drag", PressBesideTitlesIsNotATabDrag);
        yield return ("a drag that moved the window is not a tab drag", MovedWindowIsNotATabDrag);
        yield return ("a click on a tab title without dragging releases nothing", ClickWithoutDragReleasesNothing);
        yield return ("a single-tab window cannot start a tab drag", SingleTabWindowCannotTearOff);
        yield return ("a tab title that was hit-tested late is still recognised at the drop", LateTitleHitTestIsCheckedAtDrop);
        yield return ("a tear-off claims only a shown single-tab window inside the grace period", OnlyShownSingleTabWindowsInsideGrace);
        yield return ("no window is claimed until the source window has lost its tab", NothingIsClaimedUntilSourceLosesTab);
        yield return ("a window already on screen when the drag began is never claimed", WindowShownBeforeDragIsNotClaimed);
        yield return ("a drop stays pending until a window is claimed or the grace period ends", DropStaysPendingUntilClaimOrGrace);
        yield return ("a window that replaced the claimed handle is not released", ReplacedWindowIsNotReleased);
        yield return ("a window without a known tab row borrows the source window's row", UnknownRowBorrowsSourceRow);
        yield return ("a drop with no known tab row anywhere keeps merging", NoKnownRowKeepsMerging);
        yield return ("disabled tab drag detection never releases windows", DisabledDetectionNeverReleases);
        yield return ("the window released for a tab drag is never concealed as a merge source", ReleasedWindowIsNotConcealed);
        yield return ("a concealed frame Explorer shows for a torn-off tab is restored and protected", ConcealedFrameShownForTearOffIsRestored);
        yield return ("a registered window released for a tab drag is never merged", RegisteredWindowIsNotMerged);
        yield return ("a window concealed before the drop is confirmed is shown again on its own", ConcealedWindowIsShownOnceDropIsConfirmed);
        yield return ("a registration that starts before the drop is confirmed waits for it", RegistrationWaitsForDropConfirmation);
    }

    private static (FakeTabTearOffEnvironment Screen, ExplorerTabTearOffTracker Tracker) CreateTracker(int sourceTabs = 2)
    {
        var screen = new FakeTabTearOffEnvironment();
        screen.AddWindow(Source, SourceBounds, sourceTabs);
        return (screen, new ExplorerTabTearOffTracker(screen));
    }

    /// <summary>Presses on a tab title and releases at <paramref name="drop"/>; Explorer then takes the tab out of the source window.</summary>
    private static void TearOff(ExplorerTabTearOffTracker tracker, Point drop, long now = 1_000, Point? press = null,
        FakeTabTearOffEnvironment? screen = null)
    {
        tracker.HandleLeftDown(press ?? OnTab);
        tracker.HandleLeftUp(drop, now);
        if (screen != null && screen.Windows.TryGetValue(Source, out var source))
            source.TabCount = Math.Max(1, source.TabCount - 1);
    }

    private static Task TornOffTabReleasesNewWindow()
    {
        var (screen, tracker) = CreateTracker();
        TearOff(tracker, BelowTabs, screen: screen);
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);
        screen.AddWindow(Other, OtherBounds, 1);

        Check.That(tracker.TryClaim(NewWindow, 1_400), "The window Explorer opens for the dragged tab must be released.");
        Check.That(tracker.TryClaim(NewWindow, 2_500), "The answer for the released window must not change between merge checkpoints.");
        Check.That(!tracker.TryClaim(Other, 1_500), "Only one window comes out of one drag; any other window is merged as usual.");
        Check.That(!tracker.TryClaim(Source, 1_500), "The window the tab was dragged out of is never the outcome of the drag.");
        return Task.CompletedTask;
    }

    private static Task DropOnOwnRowIsNotATearOff()
    {
        var (screen, tracker) = CreateTracker(sourceTabs: 3);

        TearOff(tracker, OnSecondTab, screen: screen);
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);
        Check.That(!tracker.TryClaim(NewWindow, 1_200), "A drop on another tab title reorders the tabs; a window opened now is unrelated.");
        TearOff(tracker, BesideTabs, 2_000, screen: screen);
        Check.That(!tracker.TryClaim(NewWindow, 2_200), "A drop beside the titles stays on the tab row; a window opened now is unrelated.");
        return Task.CompletedTask;
    }

    private static Task DropOnAnotherRowIsNotATearOff()
    {
        var (screen, tracker) = CreateTracker();
        screen.AddWindow(Other, OtherBounds, 1);

        TearOff(tracker, new Point(1500, 126), screen: screen);
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);

        Check.That(!tracker.TryClaim(NewWindow, 1_200), "A tab dropped on another window's row moves there; no new window belongs to the drag.");
        return Task.CompletedTask;
    }

    private static Task PressBesideTitlesIsNotATabDrag()
    {
        var (screen, tracker) = CreateTracker();

        TearOff(tracker, OutsideWindows, press: BesideTabs, screen: screen);
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);

        Check.That(!tracker.TryClaim(NewWindow, 1_200), "Dragging the empty part of the tab row moves the window; a new window is merged as usual.");
        return Task.CompletedTask;
    }

    private static Task MovedWindowIsNotATabDrag()
    {
        var (screen, tracker) = CreateTracker();

        tracker.HandleLeftDown(OnTab);
        screen.Windows[Source].Bounds = Rectangle.FromLTRB(150, 150, 1150, 800);
        tracker.HandleLeftUp(OutsideWindows, 1_000);
        screen.Windows[Source].TabCount = 1;
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);

        Check.That(!tracker.TryClaim(NewWindow, 1_200), "A drag that moved the window was a title-bar drag; a new window is merged as usual.");
        return Task.CompletedTask;
    }

    private static Task ClickWithoutDragReleasesNothing()
    {
        var (screen, tracker) = CreateTracker();

        TearOff(tracker, new Point(OnTab.X + 3, OnTab.Y + 3), screen: screen);
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);

        Check.That(!tracker.TryClaim(NewWindow, 1_200), "A click that stays within the drag threshold is a tab switch, not a drag.");
        return Task.CompletedTask;
    }

    private static Task SingleTabWindowCannotTearOff()
    {
        var (screen, tracker) = CreateTracker(sourceTabs: 1);

        TearOff(tracker, BelowTabs);
        screen.Windows[Source].Alive = false;
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);

        Check.That(!tracker.TryClaim(NewWindow, 1_200), "Explorer never tears the only tab out of a window; a new window is merged as usual.");
        return Task.CompletedTask;
    }

    private static Task LateTitleHitTestIsCheckedAtDrop()
    {
        var (screen, tracker) = CreateTracker();
        screen.Windows[Source].RowKnown = false;

        tracker.HandleLeftDown(OnTab);
        screen.Windows[Source].RowKnown = true;
        tracker.HandleLeftUp(BelowTabs, 1_000);
        screen.Windows[Source].TabCount = 1;
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);

        Check.That(tracker.TryClaim(NewWindow, 1_200), "The title hit-test may only become ready during the drag; the press is checked again at the drop.");
        return Task.CompletedTask;
    }

    private static Task OnlyShownSingleTabWindowsInsideGrace()
    {
        var (screen, tracker) = CreateTracker(sourceTabs: 3);
        TearOff(tracker, BelowTabs, screen: screen);
        var hidden = screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);
        hidden.Shown = false;
        screen.AddWindow(Other, OtherBounds, 2);

        Check.That(!tracker.TryClaim(NewWindow, 1_200), "A frame Explorer still keeps hidden cannot be the window it opened for the drop.");
        Check.That(!tracker.TryClaim(Other, 1_200), "A window with several tabs is not the window Explorer opened for one dragged tab.");
        hidden.Shown = true;
        Check.That(!tracker.TryClaim(NewWindow, 1_000 + ExplorerTabTearOffTracker.WindowGraceMs + 1), "A window shown long after the drop is unrelated to it.");

        hidden.Shown = false;
        TearOff(tracker, BelowTabs, 5_000, screen: screen);
        hidden.Shown = true;
        Check.That(tracker.TryClaim(NewWindow, 5_100), "A later drag releases the window that follows it.");
        return Task.CompletedTask;
    }

    private static Task NothingIsClaimedUntilSourceLosesTab()
    {
        var (screen, tracker) = CreateTracker();
        tracker.HandleLeftDown(OnTab);
        tracker.HandleLeftUp(BelowTabs, 1_000);
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);

        Check.That(!tracker.TryClaim(NewWindow, 1_100), "Until Explorer has taken the tab out of the source window, no new window can be the outcome of the drop.");
        Check.That(tracker.IsPending(1_100), "The drop still waits for its window.");

        screen.Windows[Source].TabCount = 1;
        Check.That(tracker.TryClaim(NewWindow, 1_200), "Once the source window has lost its tab the new window is the outcome of the drop.");
        Check.That(!tracker.IsPending(1_200), "A claimed drop no longer waits.");
        return Task.CompletedTask;
    }

    private static Task WindowShownBeforeDragIsNotClaimed()
    {
        var (screen, tracker) = CreateTracker();
        screen.AddWindow(Other, OtherBounds, 1);
        TearOff(tracker, BelowTabs, screen: screen);
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);

        Check.That(!tracker.TryClaim(Other, 1_200), "A window that was already on screen when the drag began cannot be the window opened for it.");
        Check.That(tracker.TryClaim(NewWindow, 1_200), "The window that appeared after the drag is the outcome of the drop.");
        return Task.CompletedTask;
    }

    private static Task DropStaysPendingUntilClaimOrGrace()
    {
        var (screen, tracker) = CreateTracker(sourceTabs: 3);
        Check.That(!tracker.IsPending(500), "Nothing is pending before any drag.");

        TearOff(tracker, BelowTabs, screen: screen);
        Check.That(tracker.IsPending(1_000 + ExplorerTabTearOffTracker.WindowGraceMs), "The drop waits for its window throughout the grace period.");
        Check.That(!tracker.IsPending(1_000 + ExplorerTabTearOffTracker.WindowGraceMs + 1), "After the grace period the drop can no longer produce a window.");

        TearOff(tracker, BelowTabs, 5_000, screen: screen);
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);
        Check.That(tracker.TryClaim(NewWindow, 5_100), "The new window is the outcome of the second drop.");
        Check.That(!tracker.IsPending(5_100), "A drop whose window has been found no longer waits.");
        return Task.CompletedTask;
    }

    private static Task ReplacedWindowIsNotReleased()
    {
        var (screen, tracker) = CreateTracker();
        TearOff(tracker, BelowTabs, screen: screen);
        var released = screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);
        Check.That(tracker.TryClaim(NewWindow, 1_200), "The window opened for the drag must be released first.");

        released.Alive = false;

        Check.That(!tracker.TryClaim(NewWindow, 1_300), "Explorer reuses handles; a later window at the released handle is merged as usual.");
        return Task.CompletedTask;
    }

    private static Task UnknownRowBorrowsSourceRow()
    {
        var (screen, tracker) = CreateTracker(sourceTabs: 3);
        screen.AddWindow(Other, OtherBounds, 1).RowKnown = false;

        TearOff(tracker, new Point(1500, 126), screen: screen);
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);
        Check.That(!tracker.TryClaim(NewWindow, 1_200), "A drop where the other window's tab row must be is a move; the row is borrowed from the source window.");

        screen.Windows[NewWindow].Shown = false;
        TearOff(tracker, new Point(1500, 400), 3_000, screen: screen);
        screen.Windows[NewWindow].Shown = true;
        Check.That(tracker.TryClaim(NewWindow, 3_200), "A drop below the borrowed row is a tear-off.");
        return Task.CompletedTask;
    }

    private static Task NoKnownRowKeepsMerging()
    {
        var (screen, tracker) = CreateTracker();
        screen.AddWindow(Other, OtherBounds, 1).RowKnown = false;

        tracker.HandleLeftDown(OnTab);
        screen.Windows[Source].RowKnown = false;
        tracker.HandleLeftUp(new Point(1500, 400), 1_000);
        screen.Windows[Source].TabCount = 1;
        screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);

        Check.That(!tracker.TryClaim(NewWindow, 1_200), "Without any known tab row the drop cannot be placed, so the window is merged as before.");
        return Task.CompletedTask;
    }

    private static Task DisabledDetectionNeverReleases()
    {
        var (screen, tracker) = CreateTracker(sourceTabs: 4);

        screen.IsEnabled = false;
        TearOff(tracker, BelowTabs, screen: screen);
        var window = screen.AddWindow(NewWindow, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);
        Check.That(!tracker.TryClaim(NewWindow, 1_200), "Detection that is switched off must not release windows.");

        screen.IsEnabled = true;
        window.Shown = false;
        tracker.HandleLeftDown(OnTab);
        screen.IsEnabled = false;
        tracker.HandleLeftUp(BelowTabs, 2_000);
        screen.Windows[Source].TabCount--;
        window.Shown = true;
        Check.That(!tracker.TryClaim(NewWindow, 2_200), "A drag that ends after detection was switched off releases nothing.");

        screen.IsEnabled = true;
        window.Shown = false;
        TearOff(tracker, BelowTabs, 3_000, screen: screen);
        window.Shown = true;
        screen.IsEnabled = false;
        Check.That(!tracker.TryClaim(NewWindow, 3_200), "A claim is never made while detection is switched off.");
        return Task.CompletedTask;
    }

    /// <summary>
    /// Describes the drag of a tab out of a source window so that <paramref name="frame"/> is its outcome. With
    /// <paramref name="confirmed"/> false, Explorer has not taken the tab out of the source window yet.
    /// </summary>
    private static void TearOffOnto(Fixture fixture, nint frame, bool confirmed = true)
    {
        var source = fixture.TearOffScreen.AddWindow(Source, SourceBounds, 2);
        TearOff(fixture.TearOffTracker, BelowTabs, Environment.TickCount64);
        fixture.TearOffScreen.AddWindow(frame, Rectangle.FromLTRB(200, 400, 1200, 1050), 1);
        if (confirmed)
            source.TabCount = 1;
    }

    /// <summary>Explorer takes the dragged tab out of the source window.</summary>
    private static void ConfirmTearOff(Fixture fixture) => fixture.TearOffScreen.Windows[Source].TabCount = 1;

    private static Task ReleasedWindowIsNotConcealed() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        fixture.EnableMerging();
        try
        {
            TearOffOnto(fixture, frame.Handle);

            Check.That(!(bool)fixture.Invoke("TryHideIncomingExplorerWindow", frame.Handle)!,
                "The window Explorer opened for the dragged tab must not be taken as a merge source.");
            Check.That(!ExplorerWindowVisibility.Contains(frame.Handle), "The released window must stay on screen.");
            Check.That((bool)fixture.Invoke("IsWindowProtected", frame.Handle)!, "The released window must be protected like a Ctrl + Shift window.");
            Check.That(!(bool)fixture.Invoke("TryHideIncomingExplorerWindow", frame.Handle)!,
                "Later show events for the released window must not conceal it either.");
            Check.That(!ExplorerWindowVisibility.Contains(frame.Handle), "The released window must remain on screen after later events.");
        }
        finally
        {
            fixture.Invoke("RecoverHiddenExplorerWindows", "test-tear-off-cleanup");
        }
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task ConcealedFrameShownForTearOffIsRestored() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: false, explorerClass: true);
        fixture.EnableMerging();
        try
        {
            fixture.Invoke("HideMergeSourceWindow", frame.Handle);
            Check.That(ExplorerWindowVisibility.Contains(frame.Handle), "The preloaded frame is concealed before Explorer shows it.");

            // Explorer reuses the preloaded frame for the dragged tab and shows it after the drop.
            WinApi.ShowWindow(frame.Handle, WinApi.SW_SHOWNOACTIVATE);
            TearOffOnto(fixture, frame.Handle);

            Check.That(!(bool)fixture.Invoke("TryHideIncomingExplorerWindow", frame.Handle)!,
                "The shown frame is the window Explorer opened for the dragged tab and must not be concealed again.");
            Check.That(!ExplorerWindowVisibility.Contains(frame.Handle), "The concealed frame must be shown again for the user.");
            Check.That((WinApi.GetWindowLong(frame.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0,
                "The released frame must be fully visible again.");
            Check.Equal(0, fixture.MergeSourceCount, "The released frame must no longer be a pending merge source.");
            Check.That((bool)fixture.Invoke("IsWindowProtected", frame.Handle)!, "The released frame must be protected from later concealment.");
        }
        finally
        {
            fixture.Invoke("RecoverHiddenExplorerWindows", "test-tear-off-frame-cleanup");
        }
        return Task.CompletedTask;
    }, tabCount: 1);

    private static Task RegisteredWindowIsNotMerged() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        var concealedDuringRead = false;
        var browser = fixture.AddBrowser(out var info, handle: frame.Handle,
            readLocation: () => concealedDuringRead |= ExplorerWindowVisibility.Contains(frame.Handle));
        info.Location = null;
        fixture.SetCatalog(browser);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        try
        {
            TearOffOnto(fixture, frame.Handle);

            await (Task)fixture.Invoke("RunShellWorkAsync", (Func<Task>)(() =>
                (Task)fixture.Invoke("ProcessRegisteredShellWindowAsync", browser, info)!))!;

            Check.That(!concealedDuringRead, "Registration of the released window must not conceal it while reading its folder.");
            Check.That(info.EventsHooked && !info.Closed, "The released window must be tracked as an independent window.");
            Check.That(frame.IsAlive && frame.CloseCommandCount == 0, "The released window must never be closed by a merge.");
            Check.That(!ExplorerWindowVisibility.Contains(frame.Handle) && fixture.MergeSourceCount == 0,
                "The released window must not be left concealed or pending.");
            Check.That((bool)fixture.Invoke("IsWindowProtected", frame.Handle)!, "The released window must stay protected after registration.");
        }
        finally
        {
            fixture.Invoke("RecoverHiddenExplorerWindows", "test-tear-off-registration-cleanup");
        }
    }, tabCount: 1);

    private static Task ConcealedWindowIsShownOnceDropIsConfirmed() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        fixture.EnableMerging();
        try
        {
            TearOffOnto(fixture, frame.Handle, confirmed: false);

            Check.That((bool)fixture.Invoke("TryHideIncomingExplorerWindow", frame.Handle)!,
                "Until Explorer has moved the tab, a new window is concealed like any other so nothing can flash.");
            Check.That(ExplorerWindowVisibility.Contains(frame.Handle), "The window must be concealed while the drop is unconfirmed.");

            ConfirmTearOff(fixture);
            var shown = await Helper.DoUntilConditionAsync(() => !ExplorerWindowVisibility.Contains(frame.Handle), visible => visible, 1_500, 25);

            Check.That(shown, "Once Explorer has taken the tab out of the source window, the concealed window must be shown again without waiting for registration.");
            Check.That((WinApi.GetWindowLong(frame.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0, "The released window must be fully visible again.");
            Check.Equal(0, fixture.MergeSourceCount, "The released window must no longer be a pending merge source.");
            Check.That((bool)fixture.Invoke("IsWindowProtected", frame.Handle)!, "The released window must be protected from later concealment.");
        }
        finally
        {
            fixture.Invoke("RecoverHiddenExplorerWindows", "test-tear-off-pending-cleanup");
        }
    }, tabCount: 1);

    private static Task RegistrationWaitsForDropConfirmation() => ExplorerTabLifetimeTests.WithFixture(async fixture =>
    {
        using var frame = new RemoteExplorerFrame(visible: true, explorerClass: true);
        // Another window on screen would be the merge target; the registration must not merge into it while the drop is pending.
        using var target = new RemoteExplorerFrame(visible: true, explorerClass: true);
        var browser = fixture.AddBrowser(out var info, handle: frame.Handle);
        info.Location = null;
        fixture.SetCatalog(browser);
        fixture.SetExplorerWindows(target.Handle);
        fixture.MarkShellConnected();
        fixture.EnableMerging();
        try
        {
            TearOffOnto(fixture, frame.Handle, confirmed: false);
            var registration = (Task)fixture.Invoke("RunShellWorkAsync", (Func<Task>)(() =>
                (Task)fixture.Invoke("ProcessRegisteredShellWindowAsync", browser, info)!))!;
            await Task.Delay(300);
            Check.That(!registration.IsCompleted, "Registration must wait while the drop may still produce its window.");
            Check.That(frame.IsAlive && frame.CloseCommandCount == 0, "The waiting registration must not close the window.");

            ConfirmTearOff(fixture);
            await registration.WaitAsync(TimeSpan.FromSeconds(5));

            Check.That(info.EventsHooked && !info.Closed, "The released window must be tracked as an independent window.");
            Check.That(frame.IsAlive && frame.CloseCommandCount == 0, "The released window must never be closed by a merge.");
            Check.Equal(1, ExplorerWindowDiscovery.GetAllExplorerTabs(target.Handle).Count(), "No tab may be opened in the other window for a torn-off tab.");
            Check.That(!ExplorerWindowVisibility.Contains(frame.Handle) && fixture.MergeSourceCount == 0, "The released window must not be left concealed or pending.");
            Check.That((bool)fixture.Invoke("IsWindowProtected", frame.Handle)!, "The released window must stay protected after registration.");
        }
        finally
        {
            fixture.Invoke("RecoverHiddenExplorerWindows", "test-tear-off-wait-cleanup");
        }
    }, tabCount: 1);
}

/// <summary>A described screen: Explorer windows with bounds, tab titles, a tab row and a life that a test can end.</summary>
internal sealed class FakeTabTearOffEnvironment : IExplorerTabTearOffEnvironment
{
    private const int TabWidth = 240;
    private const int TabHeight = 32;

    internal sealed class FakeWindow(Rectangle bounds, int tabCount)
    {
        public Rectangle Bounds { get; set; } = bounds;
        public int TabCount { get; set; } = tabCount;
        /// <summary>Whether the tab-title hit-test has been computed for this window.</summary>
        public bool RowKnown { get; set; } = true;
        public bool Shown { get; set; } = true;
        public bool Alive { get; set; } = true;

        public IEnumerable<Rectangle> Tabs =>
            Enumerable.Range(0, TabCount).Select(index => new Rectangle(Bounds.Left + 10 + index * (TabWidth + 10), Bounds.Top + 10, TabWidth, TabHeight));

        public Rectangle Row => RowKnown ? Rectangle.FromLTRB(Bounds.Left, Bounds.Top + 10, Bounds.Right, Bounds.Top + 10 + TabHeight) : Rectangle.Empty;
    }

    public Dictionary<nint, FakeWindow> Windows { get; } = new();
    public bool IsEnabled { get; set; } = true;
    public int DragWidth => 4;
    public int DragHeight => 4;

    public FakeWindow AddWindow(nint handle, Rectangle bounds, int tabCount) => Windows[handle] = new FakeWindow(bounds, tabCount);

    public nint ResolveExplorerWindow(Point point) =>
        Windows.FirstOrDefault(pair => pair.Value.Alive && pair.Value.Bounds.Contains(point)).Key;

    public bool IsPointOnTab(Point point, nint explorerWindow) =>
        Windows.TryGetValue(explorerWindow, out var window) && window.RowKnown && window.Tabs.Any(tab => tab.Contains(point));

    public Rectangle GetTabRow(nint explorerWindow) =>
        Windows.TryGetValue(explorerWindow, out var window) && window.Alive ? window.Row : Rectangle.Empty;

    public int CountTabs(nint explorerWindow) => Windows.TryGetValue(explorerWindow, out var window) && window.Alive ? window.TabCount : 0;

    public bool IsWindowShown(nint explorerWindow) => Windows.TryGetValue(explorerWindow, out var window) && window.Alive && window.Shown;

    public IEnumerable<nint> ListShownExplorerWindows() => Windows.Where(pair => pair.Value.Alive && pair.Value.Shown).Select(pair => pair.Key);

    public Rectangle GetWindowBounds(nint explorerWindow) =>
        Windows.TryGetValue(explorerWindow, out var window) && window.Alive ? window.Bounds : Rectangle.Empty;

    public Func<bool> TrackIdentity(nint explorerWindow)
    {
        var window = Windows[explorerWindow];
        return () => window.Alive;
    }
}
