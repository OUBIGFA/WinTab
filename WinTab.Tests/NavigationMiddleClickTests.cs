using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class NavigationMiddleClickTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("navigation middle-click selects the unique native/UIA addition exactly once", SelectsUniqueAddition);
        yield return ("navigation middle-click waits for UIA publication after native creation", WaitsForAutomation);
        yield return ("navigation middle-click never guesses when two tabs arrive", RejectsMultipleAdditions);
        yield return ("navigation middle-click refuses replaced or missing old tabs", RejectsReplacedTabs);
        yield return ("navigation middle-click refuses changed UIA identities", RejectsReplacedAutomation);
        yield return ("navigation middle-click skips selection when Explorer already activated the new tab", AlreadyActive);
        yield return ("navigation middle-click cancels when the user selects another old tab", UserSwitchedTab);
        yield return ("navigation middle-click cancels after a slow observation returns", CancellationAfterObservation);
        yield return ("navigation middle-click does not report success without active-tab evidence", VerifiesSelection);
        yield return ("navigation middle-click reports a rejected selection", RejectedSelection);
        yield return ("navigation middle-click stops when its window identity or foreground expires", InvalidContext);
        yield return ("navigation middle-click refuses a candidate that changes during observation", ChangingCandidate);
        yield return ("navigation middle-click does not guess for duplicate identities", DuplicateIdentities);
        yield return ("navigation click gate pairs down/up without consuming the native click", PairsClick);
        yield return ("navigation click gate quarantines overlapping and late requests", OverlappingClicks);
        yield return ("navigation click gate cancels drag even when the pointer returns", CancelsDrag);
        yield return ("navigation click gate releases the worker after cancellation and disposal", CancelsLifetime);
        yield return ("navigation click gate allows the next click immediately after success", NextSuccessfulClick);
    }

    private sealed class Fixture : IDisposable
    {
        public readonly NavigationTabSet Before = new([11, 22], ["old-a", "old-b"]);
        public NavigationTabObservation Current = new(new NavigationTabSet([11, 22, 33], ["old-a", "old-b", "new-c"]), 11);
        public readonly CancellationTokenSource Lifetime = new();
        public Func<NavigationTabObservation>? OnObserve;
        public bool Valid = true, Accept = true, Activate = true;
        public int SelectCalls, Observations;
        public string? SelectedId;
        public Task<NavigationActivationResult> Run(int timeoutMs = 800) => NavigationTabActivation.RunAsync(Before, 11,
            () => Valid, () => { Observations++; return OnObserve?.Invoke() ?? Current; },
            id =>
            {
                SelectCalls++;
                SelectedId = id;
                if (Activate && Accept) Current = Current with { ActiveTab = 33 };
                return Accept;
            }, Lifetime.Token, timeoutMs, pollMs: 2);
        public void Dispose() => Lifetime.Dispose();
    }

    private static async Task SelectsUniqueAddition()
    {
        using var fixture = new Fixture();
        Check.Equal(NavigationActivationResult.Activated, await fixture.Run(), "The target must become active.");
        Check.Equal("new-c", fixture.SelectedId!, "Selection must use the new identity, not a title or index.");
        Check.Equal(1, fixture.SelectCalls, "Only the target is selected, once.");
    }

    private static async Task WaitsForAutomation()
    {
        using var fixture = new Fixture();
        fixture.OnObserve = () => fixture.Observations < 3
            ? new NavigationTabObservation(new NavigationTabSet([11, 22, 33], ["old-a", "old-b"]), 11) : fixture.Current;
        Check.Equal(NavigationActivationResult.Activated, await fixture.Run(), "UIA publication may trail HWND creation.");
        Check.Equal(1, fixture.SelectCalls, "A partial snapshot must not select anything.");
    }

    private static async Task RejectsMultipleAdditions()
    {
        using var fixture = new Fixture();
        fixture.Current = new(new NavigationTabSet([11, 22, 33, 44], ["old-a", "old-b", "new-c", "new-d"]), 11);
        Check.Equal(NavigationActivationResult.Ambiguous, await fixture.Run(), "Two new tabs cannot be attributed to one click.");
        Check.Equal(0, fixture.SelectCalls, "Do not guess the first or last addition.");
    }

    private static async Task RejectsReplacedTabs()
    {
        using var fixture = new Fixture();
        fixture.Current = new(new NavigationTabSet([11, 33], ["old-a", "new-c"]), 11);
        Check.Equal(NavigationActivationResult.Ambiguous, await fixture.Run(), "Closing an old tab invalidates the snapshot.");
        Check.Equal(0, fixture.SelectCalls, "No selection after an old tab disappears.");
    }

    private static async Task RejectsReplacedAutomation()
    {
        using var fixture = new Fixture();
        fixture.Current = new(new NavigationTabSet([11, 22, 33], ["old-a", "replacement-b", "new-c"]), 11);
        Check.That(await fixture.Run(100) != NavigationActivationResult.Activated, "UIA rebuild is not a unique addition.");
        Check.Equal(0, fixture.SelectCalls, "No selection using rebuilt identities.");
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
        using var fixture = new Fixture { Accept = false };
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
        fixture.OnObserve = () => fixture.Observations == 1 ? fixture.Current :
            new NavigationTabObservation(new NavigationTabSet([11, 22, 44], ["old-a", "old-b", "new-d"]), 11);
        Check.Equal(NavigationActivationResult.Ambiguous, await fixture.Run(), "A replacement candidate is not the original click's tab.");
        Check.Equal(0, fixture.SelectCalls, "Never follow a changing candidate.");
    }

    private static Task DuplicateIdentities()
    {
        Check.That(!NavigationTabActivation.TryGetOnlyAddition<string>(new[] { "a", "b" }, new[] { "a", "b", "b" }, out _),
            "Duplicate identities must not be interpreted as an addition.");
        Check.That(!NavigationTabActivation.TryGetOnlyAddition<string>(new[] { "a", "a" }, new[] { "a", "b", "c" }, out _),
            "A malformed baseline cannot correlate tabs.");
        return Task.CompletedTask;
    }

    private static Task PairsClick()
    {
        using var gate = new NavigationClickGate();
        var lease = gate.Begin(new Point(100, 200), 10)!;
        Check.That(!lease.Released.Task.IsCompleted, "No activation before the matching button-up.");
        gate.Release(new Point(102, 201), 30, 4, 4);
        Check.That(lease.Released.Task.IsCompletedSuccessfully && gate.IsCurrent(lease, 30), "A small click movement is permitted.");
        gate.Complete(lease, true);
        return Task.CompletedTask;
    }

    private static Task OverlappingClicks()
    {
        using var gate = new NavigationClickGate();
        var first = gate.Begin(new Point(1, 1), 100)!;
        Check.That(gate.Begin(new Point(2, 2), 120) == null && first.Token.IsCancellationRequested,
            "A second click invalidates the unresolved first request instead of stealing its late tab.");
        gate.Complete(first, false);
        Check.That(gate.Begin(new Point(2, 2), 500) == null, "A late first tab must not be assigned to another request.");
        var later = gate.Begin(new Point(2, 2), 2_101);
        Check.That(later != null, "Quarantine is bounded.");
        gate.Complete(later!, false);
        return Task.CompletedTask;
    }

    private static Task CancelsDrag()
    {
        using var gate = new NavigationClickGate();
        var lease = gate.Begin(new Point(100, 100), 0)!;
        gate.Move(new Point(110, 100), 4, 4);
        gate.Release(new Point(100, 100), 50, 4, 4);
        Check.That(!gate.IsCurrent(lease, 50), "Returning to the down point cannot turn a drag into a click.");
        gate.Complete(lease, false);
        return Task.CompletedTask;
    }

    private static Task CancelsLifetime()
    {
        var gate = new NavigationClickGate();
        var lease = gate.Begin(Point.Empty, 0)!;
        gate.Dispose();
        Check.That(lease.Token.IsCancellationRequested && !gate.IsCurrent(lease, 1), "Disposal cancels all pending work.");
        Check.That(gate.Begin(Point.Empty, 3_000) == null, "A disposed gate cannot accept new input.");
        gate.Complete(lease, false);
        return Task.CompletedTask;
    }

    private static Task NextSuccessfulClick()
    {
        using var gate = new NavigationClickGate();
        var first = gate.Begin(Point.Empty, 0)!;
        gate.Complete(first, true);
        var next = gate.Begin(Point.Empty, 20);
        Check.That(next != null, "Resolved clicks should not incur the ambiguity cooldown.");
        gate.Complete(next!, true);
        return Task.CompletedTask;
    }
}
