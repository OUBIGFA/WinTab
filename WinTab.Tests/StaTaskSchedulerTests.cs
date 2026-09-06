using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using WinTab.Helpers;

internal static class StaTaskSchedulerTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("STA shutdown does not wait forever for an unresponsive operation", ShutdownIsBounded);
        yield return ("STA worker pumps shell notification messages while idle", PumpsMessages);
    }

    private static async Task PumpsMessages()
    {
        using var scheduler = new StaTaskScheduler();
        var delivered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await Task.Factory.StartNew(() =>
        {
            Dispatcher.CurrentDispatcher.BeginInvoke(new Action(() => delivered.TrySetResult()));
        }, CancellationToken.None, TaskCreationOptions.None, scheduler);
        Check.That(await Task.WhenAny(delivered.Task, Task.Delay(500)) == delivered.Task,
            "Shell callbacks posted to the STA message loop must be delivered without another task arriving.");
    }

    private static async Task ShutdownIsBounded()
    {
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var scheduler = new StaTaskScheduler();
        var blocked = Task.Factory.StartNew(() =>
        {
            started.Set();
            release.Wait();
        }, CancellationToken.None, TaskCreationOptions.None, scheduler);
        Check.That(started.Wait(2_000), "The test operation should start.");
        var disposal = Task.Run(scheduler.Dispose);
        var completed = false;
        try
        {
            completed = await Task.WhenAny(disposal, Task.Delay(2_800)) == disposal;
        }
        finally
        {
            release.Set();
            await Task.WhenAll(blocked, disposal).WaitAsync(TimeSpan.FromSeconds(3));
        }
        Check.That(completed, "Shutdown must return within its two-second budget even when an operation is stuck.");
    }
}
