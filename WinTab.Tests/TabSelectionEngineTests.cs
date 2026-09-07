using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class TabSelectionEngineTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("TabSelectionEngine short-circuits when the target is already active", AlreadyActiveReturnsImmediately);
        yield return ("TabSelectionEngine cycles until the target becomes active", CyclesUntilTargetActive);
        yield return ("TabSelectionEngine returns false when the target never appears", ReturnsFalseWhenTargetMissing);
        yield return ("TabSelectionEngine refuses a zero target without side effects", RefusesZeroTarget);
        yield return ("TabSelectionEngine never requests an out-of-range index", DoesNotOpenNewTabsWhileCycling);
        yield return ("TabSelectionEngine waits for the delayed active tab update without skipping", WaitsForDelayedActiveUpdate);
        yield return ("TabSelectionEngine survives transient zero active readings", SurvivesTransientZeroActive);
        yield return ("TabSelectionEngine honors the total timeout even with many tabs", HonorsTotalTimeout);
        yield return ("TabSelectionEngine does not revisit indexes during cycling", DoesNotRevisitIndexes);
        yield return ("TabSelectionEngine still converges when each Explorer step is slow", SlowExplorerStillConverges);
    }

    private static async Task AlreadyActiveReturnsImmediately()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 1);
        var start = Environment.TickCount64;
        var ok = await TabSelectionEngine.CycleToTabAsync(22, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex, totalTimeoutMs: 200);

        Check.That(ok, "Cycling must return true immediately when the requested handle is already active.");
        Check.That(Environment.TickCount64 - start < 80, "An already-active target must not pay for any cycling waits.");
        Check.Equal(0, fixture.SelectionCalls.Count, "An already-active target must never send a tab-switch command.");
    }

    private static async Task CyclesUntilTargetActive()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33, 44 }, activeIndex: 0);
        fixture.OnSelectByIndex = i => fixture.SetActiveIndex(i);

        var ok = await TabSelectionEngine.CycleToTabAsync(33, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex, totalTimeoutMs: 400);

        Check.That(ok, "Cycling must succeed when the target handle is reachable by index.");
        Check.Equal<nint>(33, fixture.GetActive(), "The active tab must end up at the requested handle.");
        Check.That(fixture.SelectionCalls.Count is > 0 and <= 4, "Cycling must drive the selector forward in bounded steps.");
    }

    private static async Task ReturnsFalseWhenTargetMissing()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 0);
        fixture.OnSelectByIndex = i => fixture.SetActiveIndex(i);

        var ok = await TabSelectionEngine.CycleToTabAsync(99, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex, totalTimeoutMs: 200, pollSleepMs: 1);

        Check.That(!ok, "Cycling must report failure when no index activates the requested handle.");
        Check.Equal(0, fixture.SelectionCalls.Count, "A missing target must not switch through unrelated tabs.");
    }

    private static async Task RefusesZeroTarget()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22 }, activeIndex: 0);
        var ok = await TabSelectionEngine.CycleToTabAsync(0, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex, totalTimeoutMs: 50);

        Check.That(!ok, "A zero target handle must never report success.");
        Check.Equal(0, fixture.SelectionCalls.Count, "A zero target must not send any selection commands.");
    }

    private static async Task DoesNotOpenNewTabsWhileCycling()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 0);
        var outOfRangeRequests = 0;
        fixture.OnSelectByIndex = i =>
        {
            if (i < 0 || i >= fixture.Tabs.Length)
            {
                outOfRangeRequests++;
                return;
            }
            fixture.SetActiveIndex(i);
        };

        await TabSelectionEngine.CycleToTabAsync(33, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex, totalTimeoutMs: 300);

        Check.Equal(0, outOfRangeRequests, "Cycling must never request an index that would create or duplicate a tab.");
        Check.Equal<nint>(33, fixture.GetActive(), "Cycling must converge on the requested handle.");
    }

    private static async Task WaitsForDelayedActiveUpdate()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 0);
        var delayCount = 0;
        fixture.OnSelectByIndex = i =>
        {
            if (i == 2)
            {
                delayCount = 3;
                return;
            }
            fixture.SetActiveIndex(i);
        };
        fixture.OnGetActive = current =>
        {
            if (delayCount > 0)
            {
                delayCount--;
                return current;
            }
            if (fixture.SelectionCalls.Count > 0 && fixture.SelectionCalls[^1] == 2)
                fixture.SetActiveIndex(2);
            return current;
        };

        var ok = await TabSelectionEngine.CycleToTabAsync(33, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex, totalTimeoutMs: 400, pollSleepMs: 1);

        Check.That(ok, "The engine must wait through delayed active tab updates and still report success.");
        Check.Equal<nint>(33, fixture.GetActive(), "After waiting through the delay, the active tab must reflect the requested handle.");
    }

    private static async Task SurvivesTransientZeroActive()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33 }, activeIndex: 0);
        var emitZeroTimes = 2;
        fixture.OnSelectByIndex = i => fixture.SetActiveIndex(i);
        fixture.OnGetActive = current =>
        {
            if (emitZeroTimes > 0)
            {
                emitZeroTimes--;
                return 0;
            }
            return current;
        };

        var ok = await TabSelectionEngine.CycleToTabAsync(33, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex, totalTimeoutMs: 300, pollSleepMs: 1);

        Check.That(ok, "Transient zero active readings must not cause the engine to give up on the target.");
        Check.Equal<nint>(33, fixture.GetActive(), "The engine must converge even after seeing temporary zero readings.");
    }

    private static async Task HonorsTotalTimeout()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33, 44, 55, 66, 77, 88, 99, 1010, 1111, 1212 }, activeIndex: 0);

        var start = Environment.TickCount64;
        var ok = await TabSelectionEngine.CycleToTabAsync(9999, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex, totalTimeoutMs: 200, pollSleepMs: 1);
        var elapsed = Environment.TickCount64 - start;

        Check.That(!ok, "Missing targets must report failure.");
        Check.That(elapsed < 400, $"Total elapsed must respect totalTimeoutMs even with many tabs; observed {elapsed}ms.");
    }

    private static async Task DoesNotRevisitIndexes()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33, 44 }, activeIndex: 0);
        fixture.OnSelectByIndex = i => fixture.SetActiveIndex(i);

        var ok = await TabSelectionEngine.CycleToTabAsync(44, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex, totalTimeoutMs: 400);

        Check.That(ok, "Cycling must converge on the requested handle.");
        Check.Equal(fixture.SelectionCalls.Count, fixture.SelectionCalls.Distinct().Count(),
            $"The engine must not request the same tab index more than once; called {string.Join(',', fixture.SelectionCalls)}.");
    }

    private static async Task SlowExplorerStillConverges()
    {
        var fixture = new TabSelectionFixture(new nint[] { 11, 22, 33, 44, 55 }, activeIndex: 0);
        var pendingIndex = -1;
        var pendingDeadline = 0L;
        const int slowDelayMs = 120;

        fixture.OnSelectByIndex = i =>
        {
            pendingIndex = i;
            pendingDeadline = Environment.TickCount64 + slowDelayMs;
        };
        fixture.OnGetActive = current =>
        {
            if (pendingIndex >= 0 && Environment.TickCount64 >= pendingDeadline)
            {
                fixture.SetActiveIndex(pendingIndex);
                pendingIndex = -1;
                current = fixture.GetActiveSnapshot();
            }
            return current;
        };

        var ok = await TabSelectionEngine.CycleToTabAsync(44, fixture.GetTabs, fixture.GetActive, fixture.SendSelectByIndex,
            totalTimeoutMs: 2_500, pollSleepMs: 5, perStepTimeoutMs: 250);

        Check.That(ok, "Engine must converge on the requested tab even when Explorer takes ~120 ms per switch.");
        Check.Equal<nint>(44, fixture.GetActive(), "Final active tab must match the requested handle.");
    }

    private sealed class TabSelectionFixture(nint[] tabs, int activeIndex)
    {
        private int _activeIndex = activeIndex;

        public nint[] Tabs { get; } = tabs;
        public List<int> SelectionCalls { get; } = new();
        public Action<int>? OnSelectByIndex { get; set; }
        public Func<nint, nint>? OnGetActive { get; set; }

        public nint[] GetTabs() => Tabs;

        public nint GetActive()
        {
            var current = GetActiveSnapshot();
            return OnGetActive != null ? OnGetActive(current) : current;
        }

        public nint GetActiveSnapshot() => _activeIndex >= 0 && _activeIndex < Tabs.Length ? Tabs[_activeIndex] : 0;

        public void SetActiveIndex(int index)
        {
            if (index >= 0 && index < Tabs.Length)
                _activeIndex = index;
        }

        public void SendSelectByIndex(int i)
        {
            SelectionCalls.Add(i);
            OnSelectByIndex?.Invoke(i);
        }
    }
}
