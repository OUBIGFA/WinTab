using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class ExplorerTabWheelSwitchTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("wheel switch moves one tab per notch in either direction", MovesByNotches);
        yield return ("wheel switch stops at the first and last tab", StopsAtEnds);
        yield return ("wheel switch ignores single tabs and unknown selections", IgnoresUnusableState);
    }

    private static Task MovesByNotches()
    {
        Check.That(ExplorerTabWheelSwitchHook.ResolveTargetIndex(1, 4, 1) == 2, "Wheel down selects the next tab.");
        Check.That(ExplorerTabWheelSwitchHook.ResolveTargetIndex(1, 4, -1) == 0, "Wheel up selects the previous tab.");
        Check.That(ExplorerTabWheelSwitchHook.ResolveTargetIndex(0, 4, 2) == 2, "Queued notches accumulate.");
        return Task.CompletedTask;
    }

    private static Task StopsAtEnds()
    {
        Check.That(ExplorerTabWheelSwitchHook.ResolveTargetIndex(0, 4, -1) is null, "The first tab stays selected.");
        Check.That(ExplorerTabWheelSwitchHook.ResolveTargetIndex(3, 4, 1) is null, "The last tab stays selected.");
        Check.That(ExplorerTabWheelSwitchHook.ResolveTargetIndex(2, 4, 5) == 3, "Overshoot clamps to the last tab.");
        return Task.CompletedTask;
    }

    private static Task IgnoresUnusableState()
    {
        Check.That(ExplorerTabWheelSwitchHook.ResolveTargetIndex(0, 1, 1) is null, "A single tab cannot switch.");
        Check.That(ExplorerTabWheelSwitchHook.ResolveTargetIndex(-1, 3, 1) is null, "An unknown selection is left alone.");
        return Task.CompletedTask;
    }
}
