using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class ExplorerTabRegistrationTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("independent Explorer registration waits for its tab handle", IndependentRegistrationWaitsForTabHandle);
        yield return ("tab registration retries a transient synchronous COM failure", TransientSynchronousFailureIsRetried);
        yield return ("tab registration retries a transient asynchronous COM failure", TransientAsynchronousFailureIsRetried);
        yield return ("tab registration does not swallow cancellation", CancellationIsNotSwallowed);
    }

    private static async Task TransientSynchronousFailureIsRetried()
    {
        var attempts = 0;
        var handle = await ExplorerTabHandleResolver.WaitAsync(() =>
        {
            if (++attempts == 1)
                throw new COMException("Explorer is still registering its first tab.");
            return Task.FromResult((nint)123);
        }, 200, 1);

        Check.Equal((nint)123, handle, "A temporary COM failure must not permanently exclude the first tab from reuse.");
        Check.Equal(2, attempts, "Registration must retry the transient failure and stop once the handle is ready.");
    }

    private static async Task TransientAsynchronousFailureIsRetried()
    {
        var attempts = 0;
        var handle = await ExplorerTabHandleResolver.WaitAsync(() =>
            ++attempts == 1
                ? Task.FromException<nint>(new COMException("Explorer has not published its tab yet."))
                : Task.FromResult((nint)456), 200, 1);

        Check.Equal((nint)456, handle, "A faulted COM query must be retried within the registration deadline.");
        Check.Equal(2, attempts, "An asynchronous transient failure must not terminate registration.");
    }

    private static async Task CancellationIsNotSwallowed()
    {
        var cancelled = false;
        try
        {
            await ExplorerTabHandleResolver.WaitAsync(
                () => Task.FromException<nint>(new OperationCanceledException()), 200, 1);
        }
        catch (OperationCanceledException)
        {
            cancelled = true;
        }

        Check.That(cancelled, "Cancellation must reach the registration owner instead of looking like a missing tab.");
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
