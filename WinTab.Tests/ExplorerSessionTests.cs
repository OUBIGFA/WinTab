using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.Models;

internal static class ExplorerSessionTests
{
    private const string Home = "shell:::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}";
    private static readonly WindowIdentity Frame = Identity(100);
    private static WindowIdentity Identity(int handle, int token = 1) => new(handle, 1, 1, token);
    private static SessionTab Tab(int handle, string location, string title = "", int token = 1) =>
        new(Identity(handle, token), location, string.IsNullOrEmpty(title) ? location : title);
    private static ExplorerSession Session(string[]? locations = null, int active = 0) => new()
    {
        Locations = locations ?? [@"C:\A", @"C:\B", @"C:\C"], ActiveTabIndex = active, OrderVerified = true
    };
    private static SessionVisualTab[] Visual(SessionTab[] tabs, int active = 0) => tabs.Select((tab, index) =>
        new SessionVisualTab("uia-" + tab.Identity.Handle, tab.Title, index == active)).ToArray();

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("session capture saves single-tab windows", SingleTabIsSaved);
        yield return ("session capture retains the complete group through per-tab shutdown", WholeWindowCloseRetainsGroup);
        yield return ("session capture drops an individually closed tab after its frame settles", IndividualCloseDropsTab);
        yield return ("session capture follows visual order instead of native z-order", VisualOrderIsIndependentOfZOrder);
        yield return ("session capture retains the active duplicate path as a separate tab", DuplicatePathsAreDistinct);
        yield return ("session capture does not invent order for ambiguous equal-title folders", AmbiguousTitlesKeepAllTabs);
        yield return ("session capture rejects stale visual titles after navigation", StaleVisualTitlesDoNotVerifyOrder);
        yield return ("session capture rejects partial accessibility and native snapshots", PartialObservationsDoNotReplaceGroup);
        yield return ("session capture follows navigation immediately before close", LastNavigationIsSaved);
        yield return ("session capture rejects a navigation from a recycled native tab", RecycledTabCannotUpdateLocation);
        yield return ("session capture rebinds a recycled handle instead of inheriting its UIA owner", RecycledTabHasNewIdentity);
        yield return ("session capture excludes a tab moved to another window", MovedTabIsNotClosed);
        yield return ("session capture does not confuse a recycled live handle with a moved tab", RecycledLiveHandleIsNotMovedTab);
        yield return ("session capture forget and clear never produce closed sessions", ClearingDoesNotSave);
        yield return ("session capture does not retain caller-owned observation arrays", ObservationArraysAreIsolated);
        yield return ("session capture does not save an old partial group after exceeding the tab limit", OverLimitInvalidatesSnapshot);
        yield return ("session restore strict mode rejects an explicit folder", StrictModeRejectsFolder);
        yield return ("session restore rejects an unresolved initial location", EmptyLocationIsRejected);
        yield return ("session restore plans a single saved tab when single-tab restore is on", SingleTabIsPlanned);
        yield return ("session restore skips a single saved tab when single-tab restore is off", SingleTabIsSkippedWhenOff);
        yield return ("session restore reuses an already matching single initial tab", MatchingSingleTabIsReused);
        yield return ("session restore folder mode reuses one match without deduplicating history", OnlyOneDuplicateIsReused);
        yield return ("session restore plans preserve saved indexes when unavailable paths are skipped", SkippedLocationsKeepSavedIndexes);
        yield return ("session restore with no usable paths keeps the initial tab", NoUsablePathsKeepsInitial);
        yield return ("session restore appends in order then closes the still-active placeholder before selecting", OrderedRestoreClosesActivePlaceholder);
        yield return ("session restore any-folder normal launch keeps the start page tab first and active", AnyFolderNormalLaunchKeepsStartPage);
        yield return ("session restore requests every tab before confirming any location", RequestsAllTabsBeforeConfirming);
        yield return ("session restore an unconfirmed location stops before closing or selecting", UnconfirmedLocationStops);
        yield return ("session restore reuses a matching placeholder as the first saved tab", MatchingFirstTabReusesPlaceholder);
        yield return ("session restore selects a reused placeholder that was the saved active tab", ReusedPlaceholderCanBeActive);
        yield return ("session restore folder mode never replaces the opened folder", FolderModeNeverClosesOpenedFolder);
        yield return ("session restore folder mode preserves the opened folder and its selection", FolderRestoreKeepsInitial);
        yield return ("session restore recreates every duplicate saved tab", RestoreKeepsDuplicates);
        yield return ("session restore selects a surviving tab when the saved active path is unavailable", MissingActiveFallsBackToFirst);
        yield return ("session restore unsupported creation leaves the initial tab alone", () => FailedCreationKeepsInitial(1));
        yield return ("session restore partial creation does not close the initial tab", () => FailedCreationKeepsInitial(2));
        yield return ("session restore a failed final selection sends no further commands", FailedSelectionStops);
        yield return ("session restore never repeats an unconfirmed placeholder close", FailedCloseIsNotRetried);
        yield return ("session restore cancellation prevents further creation and placeholder close", CancelledRestoreStops);
        yield return ("session restore user navigation stops the transaction without redirecting it", UserNavigationStopsRestore);
        yield return ("session restore a matching single tab requires no creation or close", MatchingSingleTabNeedsNoMutation);
    }

    private static Task SingleTabIsSaved()
    {
        var tracker = new ExplorerSessionTracker();
        tracker.Observe(Frame, [Tab(1, @"C:\single")], 1, null, 0);
        var session = tracker.WindowClosed(Frame)!;
        Check.That(session.Locations.SequenceEqual([@"C:\single"]), "One tab is a complete restorable window, not an empty group.");
        Check.Equal(0, session.ActiveTabIndex, "The only tab must be active.");
        Check.That(session.OrderVerified, "A singleton has an unambiguous order without accessibility.");
        Check.That(tracker.WindowClosed(Frame) == null, "A window close must be consumed exactly once.");
        return Task.CompletedTask;
    }

    private static Task WholeWindowCloseRetainsGroup()
    {
        var tracker = new ExplorerSessionTracker();
        var tabs = new[] { Tab(1, @"C:\A"), Tab(2, @"C:\B"), Tab(3, @"C:\C") };
        tracker.Observe(Frame, tabs, 2, Visual(tabs, 1), 0);
        tracker.TabClosing(Frame, tabs[0].Identity, 100);
        tracker.Observe(Frame, tabs[1..], 2, Visual(tabs[1..]), 130);
        tracker.TabClosing(Frame, tabs[1].Identity, 200);
        tracker.Observe(Frame, [tabs[2]], 3, Visual([tabs[2]]), 230);
        var session = tracker.WindowClosed(Frame)!;
        Check.That(session.Locations.SequenceEqual(tabs.Select(tab => tab.Location)), "Per-tab OnQuit must not reduce a whole-window close to its last survivor.");
        Check.Equal(1, session.ActiveTabIndex, "Shutdown's transient active tab must not replace the user's last selection.");
        return Task.CompletedTask;
    }

    private static Task IndividualCloseDropsTab()
    {
        var tracker = new ExplorerSessionTracker();
        var tabs = new[] { Tab(1, @"C:\A"), Tab(2, @"C:\B"), Tab(3, @"C:\C") };
        tracker.Observe(Frame, tabs, 2, Visual(tabs, 1), 0);
        tracker.TabClosing(Frame, tabs[1].Identity, 100);
        var survivors = new[] { tabs[0], tabs[2] };
        tracker.Observe(Frame, survivors, 3, Visual(survivors, 1), 150);
        tracker.Observe(Frame, survivors, 3, Visual(survivors, 1), 150 + ExplorerSessionTracker.CloseSettleMs);
        var session = tracker.WindowClosed(Frame)!;
        Check.That(session.Locations.SequenceEqual([@"C:\A", @"C:\C"]), "A tab removed while its window stays open must not return with the later closed window.");
        Check.Equal(1, session.ActiveTabIndex, "The settled active surviving tab must be saved.");
        return Task.CompletedTask;
    }

    private static Task VisualOrderIsIndependentOfZOrder()
    {
        var tracker = new ExplorerSessionTracker();
        var tabs = new[] { Tab(1, @"C:\A"), Tab(2, @"C:\B"), Tab(3, @"C:\C") };
        tracker.Observe(Frame, [tabs[1], tabs[0], tabs[2]], 2, Visual(tabs, 1), 0);
        tracker.Observe(Frame, [tabs[2], tabs[1], tabs[0]], 3, Visual([tabs[0], tabs[2], tabs[1]], 1), 500);
        var session = tracker.WindowClosed(Frame)!;
        Check.That(session.Locations.SequenceEqual([@"C:\A", @"C:\C", @"C:\B"]), "Native activation z-order must not override a reordered visual tab strip.");
        Check.Equal(1, session.ActiveTabIndex, "The visual position of the active native tab must be retained.");
        Check.That(session.OrderVerified, "Unique linked titles and selection establish the order.");
        return Task.CompletedTask;
    }

    private static Task DuplicatePathsAreDistinct()
    {
        var tracker = new ExplorerSessionTracker();
        var tabs = new[] { Tab(1, @"C:\same"), Tab(2, @"C:\other"), Tab(3, @"C:\same") };
        tracker.Observe(Frame, tabs, 3, Visual(tabs, 2), 0);
        var session = tracker.WindowClosed(Frame)!;
        Check.That(session.Locations.SequenceEqual([@"C:\same", @"C:\other", @"C:\same"]), "Native tabs are distinct even when paths match.");
        Check.Equal(2, session.ActiveTabIndex, "The active duplicate must not collapse into the first matching path.");
        return Task.CompletedTask;
    }

    private static Task AmbiguousTitlesKeepAllTabs()
    {
        var tracker = new ExplorerSessionTracker();
        var tabs = new[] { Tab(1, @"C:\A\same", "same"), Tab(2, @"C:\B\same", "same"), Tab(3, @"C:\C\same", "same") };
        tracker.Observe(Frame, [tabs[1], tabs[0], tabs[2]], 3, Visual(tabs, 2), 0);
        var session = tracker.WindowClosed(Frame)!;
        Check.That(session.Locations.ToHashSet().SetEquals(tabs.Select(tab => tab.Location)), "Ambiguous titles must never cause paths to disappear.");
        Check.That(!session.OrderVerified, "An incomplete association must not claim a verified visual order.");
        Check.Equal(tabs[2].Location, session.Locations[session.ActiveTabIndex], "The active tab's path is still known exactly.");
        return Task.CompletedTask;
    }

    private static Task StaleVisualTitlesDoNotVerifyOrder()
    {
        var tracker = new ExplorerSessionTracker();
        var a = Tab(1, @"C:\A\same", "same");
        var b = Tab(2, @"C:\B\same", "same");
        var c = Tab(3, @"C:\C", "C");
        var oldVisual = Visual([b, a, c], 2);
        tracker.Observe(Frame, [a, b, c], 3, oldVisual, 0);
        b = b with { Location = @"C:\B\new", Title = "new" };
        tracker.Observe(Frame, [a, b, c], 3, oldVisual, 10);
        var stale = tracker.WindowClosed(Frame)!;
        Check.That(!stale.OrderVerified, "Old UIA titles must not prove a new location's visual position.");
        Check.That(stale.Locations.ToHashSet().SetEquals([a.Location, b.Location, c.Location]),
            "The navigation must still update the saved paths without losing a tab.");
        tracker.Observe(Frame, [a, b, c], 3, Visual([b, a, c], 2), 20);
        var refreshed = tracker.WindowClosed(Frame)!;
        Check.That(refreshed.OrderVerified && refreshed.Locations.SequenceEqual([b.Location, a.Location, c.Location]),
            "A matching fresh UIA snapshot may establish the actual visual order.");
        return Task.CompletedTask;
    }

    private static Task PartialObservationsDoNotReplaceGroup()
    {
        var tracker = new ExplorerSessionTracker();
        var tabs = new[] { Tab(1, @"C:\A"), Tab(2, @"C:\B") };
        tracker.Observe(Frame, tabs, 1, Visual(tabs), 0);
        tracker.Observe(Frame, tabs, 1, Visual([tabs[0]]), 1_000);
        tracker.Observe(Frame, [tabs[0], tabs[0]], 1, null, 2_000);
        tracker.Observe(Frame, tabs, 99, null, 3_000);
        Check.That(tracker.WindowClosed(Frame)!.Locations.SequenceEqual([@"C:\A", @"C:\B"]), "Partial UIA, torn walks and unknown selection must preserve the complete snapshot.");
        return Task.CompletedTask;
    }

    private static Task LastNavigationIsSaved()
    {
        var tracker = new ExplorerSessionTracker();
        tracker.Observe(Frame, [Tab(1, @"C:\before")], 1, null, 0);
        tracker.UpdateLocation(Frame, Identity(1), @"C:\after");
        tracker.TabClosing(Frame, Identity(1), 1);
        Check.Equal(@"C:\after", tracker.WindowClosed(Frame)!.Locations[0], "NavigateComplete immediately followed by OnQuit must retain the new path without another polling tick.");
        return Task.CompletedTask;
    }

    private static Task RecycledTabCannotUpdateLocation()
    {
        var tracker = new ExplorerSessionTracker();
        tracker.Observe(Frame, [Tab(1, @"C:\current", token: 2)], 1, null, 0);
        tracker.UpdateLocation(Frame, Identity(1, 1), @"C:\old-event");
        Check.Equal(@"C:\current", tracker.WindowClosed(Frame)!.Locations[0], "A delayed event must not update a new tab at the same HWND.");
        return Task.CompletedTask;
    }

    private static Task RecycledTabHasNewIdentity()
    {
        var tracker = new ExplorerSessionTracker();
        var tabs = new[] { Tab(1, @"C:\old", "folder"), Tab(2, @"C:\other", "other") };
        tracker.Observe(Frame, tabs, 2, Visual(tabs, 1), 0);
        tabs[0] = Tab(1, @"C:\new", "folder", 2);
        tracker.Observe(Frame, tabs, 1, Visual(tabs), 1);
        var session = tracker.WindowClosed(Frame)!;
        Check.That(session.Locations.SequenceEqual([@"C:\new", @"C:\other"]), "A replacement handle must be linked afresh and saved, never delayed as an old closing tab.");
        Check.Equal(0, session.ActiveTabIndex, "The replacement active tab must have its actual position.");
        return Task.CompletedTask;
    }

    private static Task MovedTabIsNotClosed()
    {
        var tracker = new ExplorerSessionTracker();
        tracker.Observe(Frame, [Tab(1, @"C:\A"), Tab(2, @"C:\B")], 1, null, 0);
        var session = tracker.WindowClosed(Frame, new HashSet<WindowIdentity> { Identity(1) })!;
        Check.That(session.Locations.SequenceEqual([@"C:\B"]), "A moved native tab must not be resurrected in the source group.");
        Check.Equal(0, session.ActiveTabIndex, "If the active tab moved, an existing surviving tab must become the fallback.");
        return Task.CompletedTask;
    }

    private static Task RecycledLiveHandleIsNotMovedTab()
    {
        var tracker = new ExplorerSessionTracker();
        tracker.Observe(Frame, [Tab(1, @"C:\closed")], 1, null, 0);
        Check.Equal(@"C:\closed", tracker.WindowClosed(Frame, new HashSet<WindowIdentity> { Identity(1, 2) })!.Locations[0],
            "A live replacement at the same numeric handle must not erase a closed tab's history.");
        return Task.CompletedTask;
    }

    private static Task ClearingDoesNotSave()
    {
        var tracker = new ExplorerSessionTracker();
        tracker.Observe(Frame, [Tab(1, @"C:\A")], 1, null, 0);
        tracker.Forget(Frame);
        Check.That(tracker.WindowClosed(Frame) == null, "Forgetting a restore-in-progress window must not produce partial history.");
        tracker.Observe(Frame, [Tab(1, @"C:\A")], 1, null, 0);
        tracker.Clear();
        Check.That(tracker.WindowClosed(Frame) == null, "Disabling capture must retire all previous observations.");
        return Task.CompletedTask;
    }

    private static Task ObservationArraysAreIsolated()
    {
        var tracker = new ExplorerSessionTracker();
        var tabs = new[] { Tab(1, @"C:\original") };
        tracker.Observe(Frame, tabs, 1, null, 0);
        tabs[0] = Tab(1, @"C:\mutated");
        Check.Equal(@"C:\original", tracker.WindowClosed(Frame)!.Locations[0], "The tracker must own its saved observation.");
        return Task.CompletedTask;
    }

    private static Task OverLimitInvalidatesSnapshot()
    {
        var tracker = new ExplorerSessionTracker();
        var tabs = Enumerable.Range(1, ExplorerSession.MaxTabs + 1).Select(index => Tab(index, @"C:\tab-" + index)).ToArray();
        tracker.Observe(Frame, tabs[..^1], 1, null, 0);
        tracker.Observe(Frame, tabs, 1, null, 1);
        Check.That(tracker.WindowClosed(Frame) == null,
            "A window with more than 100 tabs must not replace the saved session with an old 100-tab snapshot.");
        return Task.CompletedTask;
    }

    private static ExplorerSessionRestorePlan? Plan(ExplorerSession session, string initial = Home, bool anyFolder = false,
        HashSet<int>? available = null, bool singleTab = true) => ExplorerSessionRestorePlan.Create(session, initial,
            ExplorerSessionLocationPolicy.IsStartPage(initial), anyFolder, singleTab,
            available ?? Enumerable.Range(0, session.Locations.Length).ToHashSet());

    private static Task StrictModeRejectsFolder()
    {
        Check.That(Plan(Session(), @"C:\requested") == null, "Opening a specific folder must not trigger strict restoration.");
        Check.That(Plan(Session(), @"C:\requested", true) != null, "Any-folder mode must allow the same explicit request.");
        return Task.CompletedTask;
    }

    private static Task EmptyLocationIsRejected()
    {
        Check.That(Plan(Session(), string.Empty, true) == null, "An unknown initial location cannot establish user intent.");
        return Task.CompletedTask;
    }

    private static Task SingleTabIsPlanned()
    {
        var plan = Plan(Session([@"C:\single"]))!;
        Check.Equal(1, plan.Tabs.Length, "A saved singleton must not be filtered out.");
        Check.That(plan.CloseInitialTab && !plan.ActivateInitialTab,
            "A normal launch may remove only its unchanged placeholder after restoring the singleton.");
        return Task.CompletedTask;
    }

    private static Task SingleTabIsSkippedWhenOff()
    {
        Check.That(Plan(Session([@"C:\single"]), singleTab: false) == null &&
            Plan(Session([@"C:\single"]), @"C:\requested", true, singleTab: false) == null,
            "With single-tab restore off, a window that had one tab is not a group to restore.");
        Check.That(Plan(Session([@"C:\A", @"C:\B"]), singleTab: false) != null,
            "A group of two tabs is restored whether or not single-tab restore is on.");
        return Task.CompletedTask;
    }

    private static Task MatchingSingleTabIsReused()
    {
        var plan = Plan(Session([Home]))!;
        Check.That(plan.Tabs.Length == 0 && plan.ActivateInitialTab && !plan.CloseInitialTab,
            "An already matching singleton must not be duplicated or closed.");
        return Task.CompletedTask;
    }

    private static Task OnlyOneDuplicateIsReused()
    {
        var plan = Plan(Session([@"C:\A", @"C:\A", @"C:\B"]), @"C:\A", true)!;
        Check.That(plan.Tabs.Select(tab => tab.Location).SequenceEqual([@"C:\A", @"C:\B"]), "Only one matching occurrence can be supplied by the user's initial tab.");
        Check.That(plan.ActivateInitialTab && !plan.CloseInitialTab && plan.SkippedCount == 0, "Reuse is not an unavailable-path skip.");
        return Task.CompletedTask;
    }

    private static Task SkippedLocationsKeepSavedIndexes()
    {
        var plan = Plan(Session(active: 2), available: [0, 2])!;
        Check.That(plan.Tabs.Select(tab => tab.SavedIndex).SequenceEqual([0, 2]), "Filtering locations must not renumber saved selection identities.");
        Check.Equal(1, plan.SkippedCount, "The unavailable path must be reported.");
        Check.Equal(2, plan.ActiveSavedIndex, "The saved active tab keeps its original index.");
        return Task.CompletedTask;
    }

    private static Task NoUsablePathsKeepsInitial()
    {
        var plan = Plan(Session(), available: [])!;
        Check.That(plan.Tabs.Length == 0 && plan.ActivateInitialTab && !plan.CloseInitialTab && plan.SkippedCount == 3,
            "An unavailable session must never cause an empty window or close the opened folder.");
        return Task.CompletedTask;
    }

    private static async Task OrderedRestoreClosesActivePlaceholder()
    {
        var environment = new RestoreEnvironment(Home);
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session(active: 1))!, environment, CancellationToken.None);
        Check.That(result.Completed && result.RestoredCount == 3, "Every planned tab must be confirmed.");
        Check.That(environment.Tabs.Values.SequenceEqual([@"C:\A", @"C:\B", @"C:\C"]), "Tabs must be restored in saved order without an extra placeholder.");
        Check.Equal(@"C:\B", environment.Tabs[environment.Active], "The saved active tab must be selected.");
        Check.That(environment.Events.SequenceEqual(["append:C:\\A", "append:C:\\B", "append:C:\\C", "confirm:2,3,4", "close:1", "select:3"]),
            "Creation is serial and confirmed; the placeholder closes while it is still active, then the saved tab is selected.");
    }

    private static async Task AnyFolderNormalLaunchKeepsStartPage()
    {
        // A launch from the taskbar opens This PC (or the configured start folder). In any-folder mode that tab
        // is what the user opened, exactly like an explicit folder: it stays first and active.
        var plan = Plan(Session(active: 2), Home, anyFolder: true)!;
        Check.That(plan.Tabs.Length == 3 && !plan.CloseInitialTab && plan.ActivateInitialTab,
            "Any-folder mode must never close the start page the launch opened.");
        var environment = new RestoreEnvironment(Home);
        var result = await ExplorerSessionRestorer.RestoreAsync(plan, environment, CancellationToken.None);
        Check.That(result.Completed && environment.Tabs.Values.SequenceEqual([Home, @"C:\A", @"C:\B", @"C:\C"]),
            "The start page stays first and every saved tab follows it.");
        Check.That(environment.Active == environment.InitialTab && environment.CloseRequests == 0,
            "The start page remains the active tab.");
        var matching = Plan(Session([@"C:\A", Home, @"C:\B"], 0), Home, anyFolder: true)!;
        Check.That(matching.Tabs.Select(tab => tab.Location).SequenceEqual([@"C:\A", @"C:\B"]) && !matching.CloseInitialTab,
            "A saved start page is supplied by the opened one instead of being duplicated.");
    }

    private static async Task RequestsAllTabsBeforeConfirming()
    {
        var environment = new RestoreEnvironment(Home);
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session(active: 0))!, environment, CancellationToken.None);
        Check.That(result.Completed, "The restore must complete.");
        Check.That(environment.Events.SequenceEqual(["append:C:\\A", "append:C:\\B", "append:C:\\C", "confirm:2,3,4", "close:1", "select:2"]),
            "Tabs are requested back to back and confirmed together, so they appear at once rather than one per navigation.");
    }

    private static async Task UnconfirmedLocationStops()
    {
        var environment = new RestoreEnvironment(Home) { RejectConfirmation = true };
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session())!, environment, CancellationToken.None);
        Check.That(!result.Completed && result.RestoredCount == 3, "The added tabs are reported, but the restore is incomplete.");
        Check.That(environment.CloseRequests == 0 && environment.Active == environment.InitialTab &&
            !environment.Events.Any(item => item.StartsWith("select", StringComparison.Ordinal)),
            "A tab that did not reach its location must stop the restore before the placeholder is closed or a tab selected.");
    }

    private static async Task MatchingFirstTabReusesPlaceholder()
    {
        var plan = Plan(Session([Home, @"C:\A", @"C:\B"], 2))!;
        Check.That(plan.Tabs.Select(tab => tab.SavedIndex).SequenceEqual([1, 2]) && !plan.CloseInitialTab && !plan.ActivateInitialTab,
            "A placeholder already at the first saved location keeps the saved order without being recreated.");
        var environment = new RestoreEnvironment(Home);
        var result = await ExplorerSessionRestorer.RestoreAsync(plan, environment, CancellationToken.None);
        Check.That(result.Completed && environment.CloseRequests == 0 &&
            environment.Tabs.Values.SequenceEqual([Home, @"C:\A", @"C:\B"]), "The reused placeholder must stay first.");
        Check.Equal(@"C:\B", environment.Tabs[environment.Active], "The saved active tab must still be selected.");
        Check.That(Plan(Session([@"C:\A", Home]))!.CloseInitialTab,
            "A matching placeholder at another saved position cannot keep the saved order and is replaced.");
    }

    private static async Task ReusedPlaceholderCanBeActive()
    {
        var plan = Plan(Session([Home, @"C:\A"], 0))!;
        Check.That(plan.ActivateInitialTab && !plan.CloseInitialTab, "The placeholder stands in for the saved active tab.");
        var environment = new RestoreEnvironment(Home);
        var result = await ExplorerSessionRestorer.RestoreAsync(plan, environment, CancellationToken.None);
        Check.That(result.Completed && environment.Active == environment.InitialTab && environment.Tabs.Count == 2,
            "The reused placeholder must remain active after the other saved tab is added.");
    }

    private static Task FolderModeNeverClosesOpenedFolder()
    {
        foreach (var plan in new[]
        {
            Plan(Session(), @"C:\requested", true)!, Plan(Session([@"C:\A", @"C:\requested"], 1), @"C:\requested", true)!,
            Plan(Session([@"C:\requested"]), @"C:\requested", true)!
        })
            Check.That(plan.ActivateInitialTab && !plan.CloseInitialTab, "A folder the user opened must stay open and active.");
        return Task.CompletedTask;
    }

    private static async Task FolderRestoreKeepsInitial()
    {
        var environment = new RestoreEnvironment(@"C:\requested") { Selection = ["selected.txt"] };
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session(), @"C:\requested", true)!, environment, CancellationToken.None);
        Check.That(result.Completed && environment.Tabs.Count == 4, "The requested folder and all saved tabs must remain.");
        Check.Equal(environment.InitialTab, environment.Active, "The requested folder must remain the active tab.");
        Check.That(environment.Selection.SequenceEqual(["selected.txt"]) && environment.CloseRequests == 0,
            "Restoration must not navigate, close or erase the selection in the user's tab.");
    }

    private static async Task RestoreKeepsDuplicates()
    {
        var environment = new RestoreEnvironment(Home);
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session([@"C:\same", @"C:\same"], 1))!, environment, CancellationToken.None);
        Check.That(result.Completed && environment.Tabs.Count == 2, "Both independently saved duplicate tabs must return.");
        Check.Equal((nint)3, environment.Active, "The second duplicate is the saved active tab.");
    }

    private static async Task MissingActiveFallsBackToFirst()
    {
        var environment = new RestoreEnvironment(Home);
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session(active: 1), available: [0, 2])!, environment, CancellationToken.None);
        Check.That(result.Completed, "Available tabs must still restore when the old active path is unavailable.");
        Check.Equal(@"C:\A", environment.Tabs[environment.Active], "The first available saved tab is the deterministic fallback.");
    }

    private static async Task FailedCreationKeepsInitial(int failAt)
    {
        var environment = new RestoreEnvironment(Home) { FailAppendAt = failAt };
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session())!, environment, CancellationToken.None);
        Check.That(!result.Completed && result.RestoredCount == failAt - 1, "The tabs added before the failure are reported.");
        Check.That(environment.Tabs.ContainsKey(1) && environment.CloseRequests == 0 && environment.Active == 1,
            "Creation failure must retain the initial tab and avoid further focus/close commands.");
    }

    private static async Task FailedSelectionStops()
    {
        var environment = new RestoreEnvironment(Home) { RejectSelection = true };
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session(active: 2))!, environment, CancellationToken.None);
        Check.That(!result.Completed && result.RestoredCount == 3, "An unverified selection must be reported as incomplete.");
        Check.That(environment.Events.Last() == "select:4" && environment.CloseRequests == 1,
            "After a failed final selection no further commands may be sent.");
    }

    private static async Task FailedCloseIsNotRetried()
    {
        var environment = new RestoreEnvironment(Home) { RejectClose = true };
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session())!, environment, CancellationToken.None);
        Check.That(!result.Completed && environment.CloseRequests == 1, "A delayed close request must not be posted again.");
        Check.That(environment.Events.Last() == "close:1" && environment.Active == environment.InitialTab,
            "An unconfirmed close must not be followed by a selection command.");
    }

    private static async Task CancelledRestoreStops()
    {
        using var cancellation = new CancellationTokenSource();
        var environment = new RestoreEnvironment(Home) { AfterAppend = cancellation.Cancel };
        await ExpectCancellation(() => ExplorerSessionRestorer.RestoreAsync(Plan(Session())!, environment, cancellation.Token));
        Check.That(environment.Tabs.Count == 2 && environment.CloseRequests == 0,
            "Cancelling after one addition must prevent the next addition and leave the initial tab untouched.");
    }

    private static async Task UserNavigationStopsRestore()
    {
        var environment = new RestoreEnvironment(Home);
        environment.AfterAppend = () => environment.Tabs[1] = @"C:\user-navigation";
        await ExpectCancellation(() => ExplorerSessionRestorer.RestoreAsync(Plan(Session())!, environment, CancellationToken.None));
        Check.Equal(@"C:\user-navigation", environment.Tabs[1], "A competing navigation must be preserved, never redirected back.");
        Check.Equal(0, environment.CloseRequests, "A changed initial tab is no longer a disposable placeholder.");
    }

    private static async Task MatchingSingleTabNeedsNoMutation()
    {
        var environment = new RestoreEnvironment(Home);
        var result = await ExplorerSessionRestorer.RestoreAsync(Plan(Session([Home]))!, environment, CancellationToken.None);
        Check.That(result.Completed && environment.Tabs.Count == 1 && environment.CloseRequests == 0,
            "A singleton already supplied by Explorer must remain exactly one tab.");
    }

    private static async Task ExpectCancellation(Func<Task> action)
    {
        try { await action(); }
        catch (OperationCanceledException) { return; }
        throw new InvalidOperationException("The restoration must expose cancellation rather than reporting completion.");
    }

    private sealed class RestoreEnvironment(string initial) : IExplorerSessionRestoreEnvironment
    {
        public nint InitialTab => 1;
        public Dictionary<nint, string> Tabs { get; } = new() { [1] = initial };
        public nint Active = 1;
        public string[] Selection = [];
        public List<string> Events { get; } = [];
        public int CloseRequests;
        public int FailAppendAt;
        public bool RejectSelection;
        public bool RejectClose;
        public bool RejectConfirmation;
        public Action? AfterAppend;
        private int _appends;
        private int _inFlight;

        public void EnsureUnchanged()
        {
            if (Tabs.TryGetValue(1, out var location) && location != initial)
                throw new OperationCanceledException("User navigation");
        }

        public async Task<nint> AppendTabAsync(string location)
        {
            Check.Equal(1, ++_inFlight, "Only one append may be in flight at a time.");
            await Task.Yield();
            _inFlight--;
            if (++_appends == FailAppendAt) return 0;
            var handle = (nint)(_appends + 1);
            Tabs.Add(handle, location);
            Events.Add("append:" + location);
            AfterAppend?.Invoke();
            return handle;
        }

        public Task<bool> ConfirmTabsAsync(IReadOnlyList<(nint Tab, string Location)> tabs)
        {
            Events.Add("confirm:" + string.Join(",", tabs.Select(tab => tab.Tab)));
            Check.That(tabs.All(tab => Tabs.TryGetValue(tab.Tab, out var location) && location == tab.Location),
                "Every created tab must be confirmed at its own location.");
            return Task.FromResult(!RejectConfirmation);
        }

        public Task<bool> SelectTabAsync(nint tab)
        {
            Events.Add("select:" + tab);
            if (RejectSelection) return Task.FromResult(false);
            Active = tab;
            return Task.FromResult(Tabs.ContainsKey(tab));
        }

        public Task<bool> CloseInitialTabAsync()
        {
            CloseRequests++;
            Events.Add("close:1");
            if (RejectClose) return Task.FromResult(false);
            Check.That(Active == 1, "The placeholder may only be closed while it is still the active tab.");
            Tabs.Remove(1);
            // Explorer activates a neighbour of the closed active tab.
            Active = Tabs.Keys.Min();
            return Task.FromResult(true);
        }
    }
}
