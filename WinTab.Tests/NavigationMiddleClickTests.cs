using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.WinAPI;

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
        yield return ("navigation middle-click keeps observing while Explorer reorders its tab windows", KeepsObservingThroughTornRead);
        yield return ("explorer tab snapshot rereads a walk that listed a tab twice", () => SnapshotRereadsTornWalk([1, 2, 3, 2]));
        yield return ("explorer tab snapshot rereads a walk that missed a tab", () => SnapshotRereadsTornWalk([1, 2]));
        yield return ("explorer tab snapshot reports a frame that keeps changing", SnapshotReportsChangingFrame);
        yield return ("navigation input observer ignores pointer movement and the wheel", IgnoresMovementAndWheel);
        yield return ("navigation click gate pairs down/up without consuming the native click", PairsClick);
        yield return ("navigation click gate lets a new click replace an unresolved one", NewClickReplacesUnresolved);
        yield return ("navigation click gate never holds back the click after a cancelled or fruitless one", NoCooldownAfterFailure);
        yield return ("navigation click gate releases the worker after cancellation and disposal", CancelsLifetime);
        yield return ("navigation click gate allows the next click immediately after success", NextSuccessfulClick);
        yield return ("navigation click gate preserves a same-window request after an old worker retires", () => RetiredClickPreservesForeground(100));
        yield return ("navigation click gate preserves a different-window request after an old worker retires", () => RetiredClickPreservesForeground(200));
        yield return ("navigation click gate still cancels a real foreground change", ForegroundChangeCancels);
    }

    private sealed class Fixture : IDisposable
    {
        public readonly nint[] Before = [11, 22];
        public NavigationTabObservation Current = new([11, 22, 33], 11);
        public readonly CancellationTokenSource Lifetime = new();
        /// <summary>Replaces the observation; returning null models a tab walk that was not a consistent snapshot.</summary>
        public Func<NavigationTabObservation?>? OnObserve;
        public Func<NavigationSelectOutcome>? OnSelect;
        public bool Valid = true, Activate = true;
        public int SelectCalls, Observations;
        public nint SelectedTab;
        public Task<NavigationActivationResult> Run(int timeoutMs = 800) => NavigationTabActivation.RunAsync(Before, 11,
            () => Valid, () => { Observations++; return OnObserve != null ? OnObserve() : Current; },
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

    private static async Task KeepsObservingThroughTornRead()
    {
        using var fixture = new Fixture();
        // Right after creating the tab Explorer moves its window in the z-order; a walk spanning that move is
        // no snapshot of the tabs. It proves nothing about competing tabs, so the click keeps observing.
        fixture.OnObserve = () => fixture.Observations == 2 ? null : fixture.Current;
        Check.Equal(NavigationActivationResult.Activated, await fixture.Run(), "A torn tab walk must not end the click as ambiguous.");
        Check.Equal((nint)33, fixture.SelectedTab, "The new tab is still the one selected.");
        Check.Equal(1, fixture.SelectCalls, "The torn walk does not cause a second selection.");
    }

    private static Task SnapshotRereadsTornWalk(nint[] torn)
    {
        nint[] settled = [1, 3, 2];
        var walks = new Queue<nint[]>([torn, settled, settled]);
        var snapshot = ExplorerWindowDiscovery.ReadStable(walks.Dequeue);
        Check.That(snapshot != null && snapshot.SequenceEqual(settled), "A torn walk is read again until two consecutive walks agree.");
        return Task.CompletedTask;
    }

    private static Task SnapshotReportsChangingFrame()
    {
        var walk = 0;
        Check.That(ExplorerWindowDiscovery.ReadStable<nint>(() => ++walk % 2 == 0 ? [1, 2] : [1, 2, 3]) == null,
            "A frame that never settles has no snapshot; it is not reported as having fewer or more tabs.");
        nint[] duplicated = [1, 2, 2];
        Check.That(ExplorerWindowDiscovery.ReadStable(() => duplicated) == null, "A walk that lists a tab twice is never a snapshot.");
        return Task.CompletedTask;
    }

    private static Task IgnoresMovementAndWheel()
    {
        // Explorer opens the folder's tab even when the pointer or the wheel moves while the wheel button is
        // pressed, so neither may end the click. They are not observed at all.
        Check.Equal<NavigationPointerKind?>(null, NavigationInputObserver.Classify(WinApi.WM_MOUSEMOVE), "Pointer movement is not part of a middle click.");
        Check.Equal<NavigationPointerKind?>(null, NavigationInputObserver.Classify(WinApi.WM_MOUSEWHEEL), "The wheel is not part of a middle click.");
        Check.Equal<NavigationPointerKind?>(null, NavigationInputObserver.Classify(WinApi.WM_MOUSEHWHEEL), "A tilted wheel is not part of a middle click.");
        Check.Equal<NavigationPointerKind?>(NavigationPointerKind.MiddleDown, NavigationInputObserver.Classify(WinApi.WM_MBUTTONDOWN));
        Check.Equal<NavigationPointerKind?>(NavigationPointerKind.MiddleUp, NavigationInputObserver.Classify(WinApi.WM_MBUTTONUP));
        Check.Equal<NavigationPointerKind?>(NavigationPointerKind.OtherDown, NavigationInputObserver.Classify(WinApi.WM_RBUTTONDOWN),
            "Another button starts a different action, such as a context menu that opens a tab.");
        return Task.CompletedTask;
    }

    private static Task PairsClick()
    {
        using var gate = new NavigationClickGate();
        var lease = gate.Begin(100, 10, out _)!;
        Check.That(!lease.Released.Task.IsCompleted, "No activation before the matching button-up.");
        gate.Release(30);
        Check.That(lease.Released.Task.IsCompletedSuccessfully && gate.IsCurrent(lease, 30), "The button-up completes the click.");
        gate.Complete(lease);
        var late = gate.Begin(100, 100, out _)!;
        gate.Release(100 + NavigationClickGate.RequestLifetimeMs);
        Check.That(!late.Released.Task.IsCompleted && !gate.IsCurrent(late, 100 + NavigationClickGate.RequestLifetimeMs),
            "A button-up after the click's lifetime ends the click instead.");
        gate.Complete(late);
        return Task.CompletedTask;
    }

    private static Task NewClickReplacesUnresolved()
    {
        using var gate = new NavigationClickGate();
        var first = gate.Begin(100, 100, out var replacedByFirst)!;
        var second = gate.Begin(100, 120, out var replacedBySecond);
        Check.That(!replacedByFirst && replacedBySecond, "Only an unresolved click is reported as replaced.");
        Check.That(second != null && first.Token.IsCancellationRequested && !gate.IsCurrent(first, 120),
            "A new click retires the unresolved one instead of being dropped; a click that opened nothing must not cost the next one its tab.");
        gate.Complete(first);
        Check.That(gate.IsCurrent(second!, 130), "The retired click's worker cannot end the new click.");
        gate.Release(140);
        Check.That(second!.Released.Task.IsCompletedSuccessfully, "The button-up belongs to the new click.");
        gate.Complete(second);
        return Task.CompletedTask;
    }

    private static Task NoCooldownAfterFailure()
    {
        using var gate = new NavigationClickGate();
        var cancelled = gate.Begin(100, 0, out _)!;
        Check.That(gate.Cancel() && !gate.Cancel(), "A cancelled click is cancelled once and is then no longer pending.");
        var afterCancel = gate.Begin(100, 30, out var replaced);
        Check.That(afterCancel != null && !replaced && gate.IsCurrent(afterCancel, 30),
            "A cancelled click must not hold back the next one, whose tab would otherwise stay in the background.");
        gate.Complete(cancelled);
        Check.That(gate.IsCurrent(afterCancel!, 40), "Completing the cancelled click later must not disturb the current one.");
        gate.Complete(afterCancel!);
        var afterFruitless = gate.Begin(100, 50, out _);
        Check.That(afterFruitless != null && gate.IsCurrent(afterFruitless, 60), "A click that opened nothing must not hold back the next one either.");
        gate.Complete(afterFruitless!);
        return Task.CompletedTask;
    }

    private static Task CancelsLifetime()
    {
        var gate = new NavigationClickGate();
        var lease = gate.Begin(100, 0, out _)!;
        gate.Dispose();
        Check.That(lease.Token.IsCancellationRequested && !gate.IsCurrent(lease, 1), "Disposal cancels all pending work.");
        Check.That(gate.Begin(100, 3_000, out _) == null, "A disposed gate cannot accept new input.");
        gate.Complete(lease);
        return Task.CompletedTask;
    }

    private static Task NextSuccessfulClick()
    {
        using var gate = new NavigationClickGate();
        var first = gate.Begin(100, 0, out _)!;
        gate.Complete(first);
        var next = gate.Begin(100, 20, out var replaced);
        Check.That(next != null && !replaced, "A resolved click is not replaced; the next click simply starts.");
        gate.Complete(next!);
        return Task.CompletedTask;
    }

    private static Task RetiredClickPreservesForeground(nint nextWindow)
    {
        using var gate = new NavigationClickGate();
        var old = gate.Begin(100, 0, out _)!;
        var next = gate.Begin(nextWindow, 20, out _)!;

        // The old worker's finally can run only after the next request has acquired its window.
        gate.Complete(old);
        Check.That(!gate.CancelIfForegroundChanged(nextWindow), "The new request still owns its foreground window after old cleanup.");
        Check.That(!next.Token.IsCancellationRequested && gate.IsCurrent(next, 30), "Old cleanup must not cancel or retire the new request.");
        gate.Release(40);
        Check.That(next.Released.Task.IsCompletedSuccessfully, "The new request must still receive its button-up.");
        gate.Complete(next);
        Check.That(!gate.CancelIfForegroundChanged(300), "A completed request must not leave a stale foreground owner.");
        return Task.CompletedTask;
    }

    private static Task ForegroundChangeCancels()
    {
        using var gate = new NavigationClickGate();
        var lease = gate.Begin(100, 0, out _)!;
        Check.That(!gate.CancelIfForegroundChanged(100), "A foreground event from the click's own window is not a cancellation.");
        Check.That(gate.CancelIfForegroundChanged(200), "Moving to another window must still cancel the request.");
        Check.That(lease.Token.IsCancellationRequested && !gate.IsCurrent(lease, 20), "No late activation may follow a real foreground change.");
        gate.Complete(lease);
        var next = gate.Begin(200, 30, out _);
        Check.That(next != null && gate.IsCurrent(next, 40), "A click in the window the user moved to starts at once.");
        gate.Complete(next!);
        return Task.CompletedTask;
    }
}

