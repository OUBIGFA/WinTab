using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Hooks;

/// <summary>
/// A middle click on the navigation pane makes Explorer open a tab in the background. WinTab records the
/// window's tabs when the button goes down and brings the one tab that appears afterwards to the front.
/// Nothing is prepared ahead of the click, so nothing can be stale when the click arrives.
/// </summary>
internal static class NavigationMiddleClickTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("navigation middle-click selects the unique new tab exactly once", SelectsUniqueAddition);
        yield return ("navigation middle-click waits until the tab strip lists the new tab", WaitsForTabStrip);
        yield return ("navigation middle-click never guesses when two tabs arrive", RejectsMultipleAdditions);
        yield return ("navigation middle-click refuses replaced or missing old tabs", RejectsReplacedTabs);
        yield return ("navigation middle-click skips selection when Explorer already activated the new tab", AlreadyActive);
        yield return ("navigation middle-click cancels when the user selects another old tab", UserSwitchedTab);
        yield return ("navigation middle-click cancels after a slow observation returns", CancellationAfterObservation);
        yield return ("navigation middle-click does not report success without active-tab evidence", VerifiesSelection);
        yield return ("navigation middle-click reports a rejected selection", RejectedSelection);
        yield return ("navigation middle-click stops when its window identity or foreground expires", InvalidContext);
        yield return ("navigation middle-click refuses a candidate that changes during observation", ChangingCandidate);
        yield return ("navigation middle-click does not guess for duplicate handles", DuplicateHandles);
        yield return ("navigation click gate pairs down/up without consuming the native click", PairsClick);
        yield return ("navigation click gate quarantines overlapping and late requests", OverlappingClicks);
        yield return ("navigation click gate cancels drag even when the pointer returns", CancelsDrag);
        yield return ("navigation click gate releases the worker after cancellation and disposal", CancelsLifetime);
        yield return ("navigation click gate allows the next click immediately after success", NextSuccessfulClick);
        yield return ("navigation click gate forgets a click beside the folder items without a cooldown", DiscardedClickHasNoCooldown);
        yield return ("navigation click gate preserves a same-window request after an old worker retires", () => RetiredClickPreservesForeground(100));
        yield return ("navigation click gate preserves a different-window request after an old worker retires", () => RetiredClickPreservesForeground(200));
        yield return ("navigation click gate still cancels a real foreground change", ForegroundChangeCancels);
    }

    private sealed class Fixture : IDisposable
    {
        public readonly nint[] Before = [11, 22];
        public NavigationTabObservation Current = new([11, 22, 33], 11);
        public readonly CancellationTokenSource Lifetime = new();
        public Func<NavigationTabObservation>? OnObserve;
        public Func<NavigationSelectOutcome>? OnSelect;
        public bool Valid = true, Activate = true;
        public int SelectCalls, Observations;
        public nint SelectedTab;
        public Task<NavigationActivationResult> Run(int timeoutMs = 800) => NavigationTabActivation.RunAsync(Before, 11,
            () => Valid, () => { Observations++; return OnObserve?.Invoke() ?? Current; },
            tab =>
            {
                SelectCalls++;
                SelectedTab = tab;
                var outcome = OnSelect?.Invoke() ?? NavigationSelectOutcome.Selected;
                if (outcome == NavigationSelectOutcome.Selected && Activate) Current = Current with { ActiveTab = tab };
                return outcome;
            }, Lifetime.Token, timeoutMs, pollMs: 2);
        public void Dispose() => Lifetime.Dispose();
    }

    private static async Task SelectsUniqueAddition()
    {
        using var fixture = new Fixture();
        Check.Equal(NavigationActivationResult.Activated, await fixture.Run(), "The target must become active.");
        Check.Equal((nint)33, fixture.SelectedTab, "Selection must name the new native tab, not an index or a title.");
        Check.Equal(1, fixture.SelectCalls, "Only the target is selected, once.");
    }

    private static async Task WaitsForTabStrip()
    {
        using var fixture = new Fixture();
        fixture.OnSelect = () => fixture.SelectCalls < 3 ? NavigationSelectOutcome.NotReady : NavigationSelectOutcome.Selected;
        Check.Equal(NavigationActivationResult.Activated, await fixture.Run(), "The tab strip may list the new tab after its window exists.");
        Check.Equal(3, fixture.SelectCalls, "Selection is retried until the strip is ready and then performed once.");
    }

    private static async Task RejectsMultipleAdditions()
    {
        using var fixture = new Fixture();
        fixture.Current = new([11, 22, 33, 44], 11);
        Check.Equal(NavigationActivationResult.Ambiguous, await fixture.Run(), "Two new tabs cannot be attributed to one click.");
        Check.Equal(0, fixture.SelectCalls, "Do not guess the first or last addition.");
    }

    private static async Task RejectsReplacedTabs()
    {
        using var fixture = new Fixture();
        fixture.Current = new([11, 33], 11);
        Check.Equal(NavigationActivationResult.Ambiguous, await fixture.Run(), "Closing an old tab invalidates the click.");
        Check.Equal(0, fixture.SelectCalls, "No selection after an old tab disappears.");
    }

    private static async Task AlreadyActive()
    {
        using var fixture = new Fixture();
        fixture.Current = fixture.Current with { ActiveTab = 33 };
        Check.Equal(NavigationActivationResult.AlreadyActive, await fixture.Run(), "Native foreground opening needs no correction.");
        Check.Equal(0, fixture.SelectCalls, "Do not reselect an active target.");
    }

    private static async Task UserSwitchedTab()
    {
        using var fixture = new Fixture();
        fixture.Current = fixture.Current with { ActiveTab = 22 };
        Check.Equal(NavigationActivationResult.Cancelled, await fixture.Run(), "Respect the user's new tab selection.");
        Check.Equal(0, fixture.SelectCalls, "Do not switch back after user input.");
    }

    private static async Task CancellationAfterObservation()
    {
        using var fixture = new Fixture();
        fixture.OnObserve = () => { fixture.Lifetime.Cancel(); return fixture.Current; };
        Check.Equal(NavigationActivationResult.Cancelled, await fixture.Run(), "Recheck cancellation after accessibility calls.");
        Check.Equal(0, fixture.SelectCalls, "A cancelled query must not cause a late action.");
    }

    private static async Task VerifiesSelection()
    {
        using var fixture = new Fixture { Activate = false };
        Check.Equal(NavigationActivationResult.TimedOut, await fixture.Run(150), "Accepted commands are not proof of selection.");
        Check.Equal(1, fixture.SelectCalls, "Do not repeatedly force a slow or rejected operation.");
    }

    private static async Task RejectedSelection()
    {
        using var fixture = new Fixture();
        fixture.OnSelect = () => NavigationSelectOutcome.Rejected;
        Check.Equal(NavigationActivationResult.SelectionRejected, await fixture.Run(), "Return the failed action outcome.");
    }

    private static async Task InvalidContext()
    {
        using var fixture = new Fixture { Valid = false };
        Check.Equal(NavigationActivationResult.Cancelled, await fixture.Run(), "An expired context cannot activate.");
        Check.Equal(0, fixture.Observations, "Do not query an expired window.");
        Check.Equal(0, fixture.SelectCalls, "Do not touch an expired window.");
    }

    private static async Task ChangingCandidate()
    {
        using var fixture = new Fixture();
        fixture.OnObserve = () => fixture.Observations == 1 ? fixture.Current : new NavigationTabObservation([11, 22, 44], 11);
        Check.Equal(NavigationActivationResult.Ambiguous, await fixture.Run(), "A replacement candidate is not the original click's tab.");
        Check.Equal(0, fixture.SelectCalls, "Never follow a changing candidate.");
    }

    private static Task DuplicateHandles()
    {
        Check.That(!NavigationTabActivation.TryGetOnlyAddition<nint>([11, 22], [11, 22, 22], out _),
            "Duplicate handles must not be interpreted as an addition.");
        Check.That(!NavigationTabActivation.TryGetOnlyAddition<nint>([11, 11], [11, 22, 33], out _),
            "A malformed baseline cannot correlate tabs.");
        return Task.CompletedTask;
    }

    private static Task PairsClick()
    {
        using var gate = new NavigationClickGate();
        var lease = gate.Begin(100, new Point(100, 200), 10)!;
        Check.That(!lease.Released.Task.IsCompleted, "No activation before the matching button-up.");
        gate.Release(new Point(102, 201), 30, 4, 4);
        Check.That(lease.Released.Task.IsCompletedSuccessfully && gate.IsCurrent(lease, 30), "A small click movement is permitted.");
        gate.Complete(lease, true);
        return Task.CompletedTask;
    }

    private static Task OverlappingClicks()
    {
        using var gate = new NavigationClickGate();
        var first = gate.Begin(100, new Point(1, 1), 100)!;
        Check.That(gate.Begin(100, new Point(2, 2), 120) == null && first.Token.IsCancellationRequested,
            "A second click invalidates the unresolved first request instead of stealing its late tab.");
        gate.Complete(first, false);
        Check.That(gate.Begin(100, new Point(2, 2), 500) == null, "A late first tab must not be assigned to another request.");
        var later = gate.Begin(100, new Point(2, 2), 2_101);
        Check.That(later != null, "Quarantine is bounded.");
        gate.Complete(later!, false);
        return Task.CompletedTask;
    }

    private static Task CancelsDrag()
    {
        using var gate = new NavigationClickGate();
        var lease = gate.Begin(100, new Point(100, 100), 0)!;
        gate.Move(new Point(110, 100), 4, 4);
        gate.Release(new Point(100, 100), 50, 4, 4);
        Check.That(!gate.IsCurrent(lease, 50), "Returning to the down point cannot turn a drag into a click.");
        gate.Complete(lease, false);
        return Task.CompletedTask;
    }

    private static Task CancelsLifetime()
    {
        var gate = new NavigationClickGate();
        var lease = gate.Begin(100, Point.Empty, 0)!;
        gate.Dispose();
        Check.That(lease.Token.IsCancellationRequested && !gate.IsCurrent(lease, 1), "Disposal cancels all pending work.");
        Check.That(gate.Begin(100, Point.Empty, 3_000) == null, "A disposed gate cannot accept new input.");
        gate.Complete(lease, false);
        return Task.CompletedTask;
    }

    private static Task NextSuccessfulClick()
    {
        using var gate = new NavigationClickGate();
        var first = gate.Begin(100, Point.Empty, 0)!;
        gate.Complete(first, true);
        var next = gate.Begin(100, Point.Empty, 20);
        Check.That(next != null, "Resolved clicks should not incur the ambiguity cooldown.");
        gate.Complete(next!, true);
        return Task.CompletedTask;
    }

    private static Task RetiredClickPreservesForeground(nint nextWindow)
    {
        using var gate = new NavigationClickGate();
        var old = gate.Begin(100, Point.Empty, 0)!;
        gate.Discard(old);
        var next = gate.Begin(nextWindow, Point.Empty, 20)!;

        // The old worker's finally can run only after the next request has acquired its window.
        gate.Complete(old, false);
        gate.Discard(old);
        Check.That(!gate.CancelIfForegroundChanged(nextWindow), "The new request still owns its foreground window after old cleanup.");
        Check.That(!next.Token.IsCancellationRequested && gate.IsCurrent(next, 30), "Old cleanup must not cancel or retire the new request.");
        gate.Release(Point.Empty, 40, 4, 4);
        Check.That(next.Released.Task.IsCompletedSuccessfully, "The new request must still receive its button-up.");
        gate.Complete(next, true);
        Check.That(!gate.CancelIfForegroundChanged(300), "A completed request must not leave a stale foreground owner.");
        return Task.CompletedTask;
    }

    private static Task ForegroundChangeCancels()
    {
        using var gate = new NavigationClickGate();
        var lease = gate.Begin(100, Point.Empty, 0)!;
        Check.That(!gate.CancelIfForegroundChanged(100), "A foreground event from the click's own window is not a cancellation.");
        Check.That(gate.CancelIfForegroundChanged(200), "Moving to another window must still cancel the request.");
        Check.That(lease.Token.IsCancellationRequested && !gate.IsCurrent(lease, 20), "No late activation may follow a real foreground change.");
        gate.Complete(lease, false);
        Check.That(gate.Begin(200, Point.Empty, 30) == null, "A cancelled click must retain its late-tab quarantine.");
        return Task.CompletedTask;
    }

    private static Task DiscardedClickHasNoCooldown()
    {
        using var gate = new NavigationClickGate();
        var beside = gate.Begin(100, new Point(5, 5), 0)!;
        gate.Discard(beside);
        Check.That(beside.Token.IsCancellationRequested && !gate.IsCurrent(beside, 1), "A forgotten click must not keep its worker alive.");
        var next = gate.Begin(100, new Point(6, 6), 20);
        Check.That(next != null, "A click beside the folder items opens nothing, so the next click must not be held back.");
        gate.Complete(beside, false);
        Check.That(gate.IsCurrent(next!, 30), "Completing the forgotten click later must not disturb the current one.");
        gate.Complete(next!, true);
        return Task.CompletedTask;
    }
}
