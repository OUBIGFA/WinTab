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
        yield return ("wheel throttle drops notches that arrive while a switch holds", DropsNotchesDuringHold);
        yield return ("low wheel sensitivity needs two notches per switch", LowNeedsTwoNotches);
        yield return ("wheel throttle discards a partial gesture after a pause or reversal", PartialGestureExpires);
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

    private static Task DropsNotchesDuringHold()
    {
        var throttle = new WheelSwitchThrottle();
        var medium = WheelSwitchSensitivity.Medium;
        Check.That(throttle.Accept(-120, 1_000, medium) == 1, "The first notch down switches immediately.");
        Check.That(throttle.Accept(-120, 1_050, medium) == 0, "A notch inside the hold is dropped, not queued.");
        Check.That(throttle.Accept(-120, 1_100, medium) == 0, "Further notches inside the hold are dropped too.");
        Check.That(throttle.Accept(120, 1_300, medium) == -1, "After the hold, wheel up switches to the left.");
        var high = new WheelSwitchThrottle();
        Check.That(high.Accept(-120, 0, WheelSwitchSensitivity.High) == 1 &&
                   high.Accept(-120, 80, WheelSwitchSensitivity.High) == 1, "High sensitivity follows quick notches.");
        return Task.CompletedTask;
    }

    private static Task LowNeedsTwoNotches()
    {
        var throttle = new WheelSwitchThrottle();
        var low = WheelSwitchSensitivity.Low;
        Check.That(throttle.Accept(-120, 1_000, low) == 0, "One notch is not enough at low sensitivity.");
        Check.That(throttle.Accept(-120, 1_100, low) == 1, "The second notch switches.");
        return Task.CompletedTask;
    }

    private static Task PartialGestureExpires()
    {
        var throttle = new WheelSwitchThrottle();
        var low = WheelSwitchSensitivity.Low;
        Check.That(throttle.Accept(-120, 1_000, low) == 0, "A partial gesture starts.");
        Check.That(throttle.Accept(-120, 2_000, low) == 0, "After a pause the earlier notch no longer counts.");
        Check.That(throttle.Accept(120, 2_100, low) == 0, "A reversal starts over as well.");
        Check.That(throttle.Accept(-40, 0, WheelSwitchSensitivity.High) == 0, "A fraction of a notch does not switch.");
        return Task.CompletedTask;
    }
}
