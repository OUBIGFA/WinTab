using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;

internal static class ExplorerTabRegistrationTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("independent Explorer registration waits for its tab handle", IndependentRegistrationWaitsForTabHandle);
    }

    private static async Task IndependentRegistrationWaitsForTabHandle()
    {
        var resolverType = typeof(WinTab.Hooks.ExplorerWatcher).Assembly
            .GetType("WinTab.Hooks.ExplorerTabHandleResolver");
        Check.That(resolverType != null, "The independent-window registration must use a testable tab-handle resolver.");

        var waitMethod = resolverType!.GetMethod(
            "WaitAsync",
            BindingFlags.Public | BindingFlags.Static,
            binder: null,
            types: [typeof(Func<Task<nint>>), typeof(int), typeof(int)],
            modifiers: null);
        Check.That(waitMethod != null, "The tab-handle resolver must expose its asynchronous wait contract.");

        var attempts = 0;
        var task = (Task<nint>)waitMethod!.Invoke(null, [
            (Func<Task<nint>>)(() => Task.FromResult(++attempts >= 3 ? (nint)123 : 0)),
            200,
            1
        ])!;

        var handle = await task;
        Check.Equal((nint)123, handle, "Registration must wait until Explorer publishes a non-zero tab handle.");
        Check.Equal(3, attempts, "Registration should stop polling as soon as the tab handle is available.");
    }
}
