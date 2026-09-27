using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WinTab.Hooks;
using WinTab.Managers;

internal static class HookManagerTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("a hook that cannot start is reported without stopping the caller", HookStartFailureIsContained);
        yield return ("a hook that cannot stop is reported without stopping the caller", HookStopFailureIsContained);
        yield return ("hook changes are applied only when the state differs", HookChangeIsIdempotent);
        yield return ("features whose hook differs from the setting are listed in order", ListsMismatchedFeatures);
    }

    private static Task ListsMismatchedFeatures()
    {
        var mismatched = HookManager.FindMismatched(
        [
            (HookFeature.MergeWindows, true, new FakeHook { IsHookActive = true }),
            (HookFeature.DoubleClickClose, true, new FakeHook()),
            (HookFeature.MiddleClickForeground, false, new FakeHook()),
            (HookFeature.WheelSwitch, false, new FakeHook { IsHookActive = true })
        ]);
        Check.Equal("DoubleClickClose,WheelSwitch", string.Join(",", mismatched),
            "An enabled hook that did not start and a disabled hook that still runs must both be listed.");
        return Task.CompletedTask;
    }

    private static Task HookStartFailureIsContained()
    {
        var hook = new FakeHook { FailStart = true };
        Check.That(!HookManager.TryChangeHookStatus(hook, true), "A failed start must be reported as a failure.");
        Check.That(!hook.IsHookActive, "A failed start must leave the hook inactive.");
        return Task.CompletedTask;
    }

    private static Task HookStopFailureIsContained()
    {
        var hook = new FakeHook { IsHookActive = true, FailStop = true };
        Check.That(!HookManager.TryChangeHookStatus(hook, false), "A failed stop must be reported as a failure.");
        return Task.CompletedTask;
    }

    private static Task HookChangeIsIdempotent()
    {
        var hook = new FakeHook();
        Check.That(HookManager.TryChangeHookStatus(hook, true) && hook.IsHookActive, "A start must activate the hook.");
        Check.That(HookManager.TryChangeHookStatus(hook, true), "Starting an active hook must succeed without starting it again.");
        Check.Equal(1, hook.Starts, "An active hook must not be started twice.");
        Check.That(HookManager.TryChangeHookStatus(hook, false) && !hook.IsHookActive, "A stop must deactivate the hook.");
        return Task.CompletedTask;
    }

    private sealed class FakeHook : IHook
    {
        public bool FailStart { get; init; }
        public bool FailStop { get; init; }
        public int Starts { get; private set; }
        public bool IsHookActive { get; set; }

        public void StartHook()
        {
            if (FailStart) throw new InvalidOperationException("The input observer could not be installed.");
            Starts++;
            IsHookActive = true;
        }

        public void StopHook()
        {
            if (FailStop) throw new InvalidOperationException("The input observer could not be removed.");
            IsHookActive = false;
        }

        public void Dispose() { }
    }
}
