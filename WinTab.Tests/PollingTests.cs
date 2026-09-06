using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;

internal static class PollingTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("cancelled polling never starts another operation", CancelledPollingDoesNotRun);
        yield return ("polling rejects a result produced after cancellation", CancellationDuringOperation);
        yield return ("polling bounds an asynchronous operation that never finishes", StalledOperationTimesOut);
    }

    private static async Task CancelledPollingDoesNotRun()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var calls = 0;
        await ExpectCancellation(Helper.DoUntilConditionAsync(() => ++calls, value => value > 0,
            cancellationToken: cancellation.Token));
        Check.Equal(0, calls, "Cancelled work must not call the operation.");
        await ExpectCancellation(Helper.DoUntilConditionAsync(() => Task.FromResult(++calls), value => value > 0,
            cancellationToken: cancellation.Token));
        Check.Equal(0, calls, "The asynchronous overload must also skip cancelled work.");
    }

    private static async Task CancellationDuringOperation()
    {
        using var cancellation = new CancellationTokenSource();
        await ExpectCancellation(Helper.DoUntilConditionAsync(() =>
        {
            cancellation.Cancel();
            return true;
        }, result => result, cancellationToken: cancellation.Token));
    }

    private static async Task StalledOperationTimesOut()
    {
        var pending = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
        var polling = Helper.DoUntilNotDefaultAsync(() => pending.Task, timeMs: 40);
        try
        {
            var finished = await Task.WhenAny(polling, Task.Delay(1_000));
            Check.That(ReferenceEquals(polling, finished), "The polling deadline must bound the entire wait.");
            try { await polling; }
            catch (TimeoutException) { return; }
            throw new InvalidOperationException("An unfinished operation must report a timeout.");
        }
        finally
        {
            pending.TrySetResult(1);
            try { await polling; }
            catch (TimeoutException) { }
        }
    }

    private static async Task ExpectCancellation(Task task)
    {
        try { await task; }
        catch (OperationCanceledException) { return; }
        throw new InvalidOperationException("Cancelled polling must report cancellation, not success.");
    }
}
