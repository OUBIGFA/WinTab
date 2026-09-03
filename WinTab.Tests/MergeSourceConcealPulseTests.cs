using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class MergeSourceConcealPulseTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("MergeSourceConcealPulse runs while enabled and stops after its duration", IsDurationBound);
        yield return ("MergeSourceConcealPulse has an absolute ceiling", HasAbsoluteCeiling);
        yield return ("MergeSourceConcealPulse never restarts itself in finally", NeverRestartsItselfInFinally);
        yield return ("MergeSourceConcealPulse sleep period is at least 25 ms", SleepIsAtLeast25Ms);
        yield return ("MergeSourceConcealPulse does not start when disabled", DoesNotStartWhenDisabled);
        yield return ("MergeSourceConcealPulse Stop ends the running pulse early", StopEndsPulseEarly);
    }

    private static async Task IsDurationBound()
    {
        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 150, sleepMs: 5);
        var callCount = 0;

        pulse.Start(() => true, () => Interlocked.Increment(ref callCount), durationMs: 40);
        await Task.Delay(80);
        var countAfterPulse = Volatile.Read(ref callCount);
        await Task.Delay(60);

        Check.That(countAfterPulse > 0, "The pulse must run at least one scan while enabled.");
        Check.Equal(countAfterPulse, Volatile.Read(ref callCount), "The pulse must not continue scanning after its window expires.");
    }

    private static async Task HasAbsoluteCeiling()
    {
        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 90, sleepMs: 5);
        var callCount = 0;

        for (var i = 0; i < 5; i++)
        {
            pulse.Start(() => true, () => Interlocked.Increment(ref callCount), durationMs: 500);
            await Task.Delay(20);
        }

        await Task.Delay(80);
        var countAfterCeiling = Volatile.Read(ref callCount);
        await Task.Delay(80);

        Check.That(countAfterCeiling > 0, "The pulse must run while it is inside the bounded window.");
        Check.Equal(countAfterCeiling, Volatile.Read(ref callCount), "A stream of Start calls must not extend the pulse beyond its absolute ceiling.");
    }

    private static async Task NeverRestartsItselfInFinally()
    {
        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 80, sleepMs: 5);
        var callCount = 0;

        pulse.Start(() => true, () => Interlocked.Increment(ref callCount), durationMs: 30);
        await Task.Delay(80);
        var countAfterExit = Volatile.Read(ref callCount);
        await Task.Delay(80);

        Check.That(countAfterExit > 0, "The pulse must run before exiting.");
        Check.Equal(countAfterExit, Volatile.Read(ref callCount), "The pulse worker must not re-arm itself from inside its own exit path.");
    }

    private static async Task SleepIsAtLeast25Ms()
    {
        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 140, sleepMs: 25);
        var callCount = 0;

        pulse.Start(() => true, () => Interlocked.Increment(ref callCount), durationMs: 100);
        await Task.Delay(140);

        Check.That(Volatile.Read(ref callCount) <= 7, "The pulse worker must not run as a tight loop while scanning Explorer windows.");
    }

    private static async Task DoesNotStartWhenDisabled()
    {
        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 200, sleepMs: 5);
        var callCount = 0;

        pulse.Start(() => false, () => Interlocked.Increment(ref callCount), durationMs: 100);
        await Task.Delay(60);

        Check.Equal(0, Volatile.Read(ref callCount), "A disabled pulse must never invoke the scan callback.");
    }

    private static async Task StopEndsPulseEarly()
    {
        var pulse = new MergeSourceConcealPulse(absoluteCeilingMs: 2_000, sleepMs: 5);
        var callCount = 0;

        pulse.Start(() => true, () => Interlocked.Increment(ref callCount), durationMs: 1_000);
        await Task.Delay(40);
        pulse.Stop();
        await Task.Delay(30);
        var countAfterStop = Volatile.Read(ref callCount);
        await Task.Delay(60);

        Check.That(countAfterStop > 0, "The pulse must have run before Stop was called.");
        Check.Equal(countAfterStop, Volatile.Read(ref callCount), "Stop must end the pulse well before its requested duration.");
    }
}
