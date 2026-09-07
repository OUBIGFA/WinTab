using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;

internal static class BackgroundWorkTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("registration storms retain only one running and one pending scan", CoalescesWork);
        yield return ("invalidated refreshes cannot publish an old result or overlap", InvalidatedRefreshDoesNotPublish);
        yield return ("forgotten refreshes cannot repopulate a new window cache", ForgottenRefreshDoesNotPublish);
        yield return ("a stalled window refresh does not block another window", SlowRefreshDoesNotBlockAnotherWindow);
        yield return ("window refresh storms keep a bounded number of active queries", RefreshConcurrencyIsBounded);
    }

    private static async Task CoalescesWork()
    {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var calls = 0;
        var worker = new CoalescingAsyncWork(async () =>
        {
            Interlocked.Increment(ref calls);
            started.TrySetResult();
            await release.Task;
        });
        worker.Request();
        await started.Task.WaitAsync(TimeSpan.FromSeconds(2));
        for (var index = 0; index < 1_000; index++)
            worker.Request();
        Check.Equal(1, Volatile.Read(ref calls), "The active scan must not overlap another scan.");
        release.SetResult();
        await worker.WhenIdle.WaitAsync(TimeSpan.FromSeconds(2));
        Check.Equal(2, calls, "A storm should produce just one follow-up scan.");
        await worker.StopAsync();
        worker.Request();
        Check.Equal(2, calls, "Stopped workers must ignore later events.");
    }

    private static async Task InvalidatedRefreshDoesNotPublish()
    {
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var calls = 0;
        using var cache = new BackgroundRefreshCache<int, string>(key =>
        {
            var current = Interlocked.Increment(ref calls);
            if (current == 1)
            {
                started.Set();
                release.Wait();
            }
            return current == 1 ? "old" : "new";
        });
        cache.Request(1);
        Check.That(started.Wait(2_000), "The refresh should start.");
        try
        {
            for (var index = 0; index < 100; index++)
                cache.Invalidate(1);
            Check.Equal(1, calls, "Invalidation must not launch concurrent automation queries.");
        }
        finally
        {
            release.Set();
        }
        await WaitForValue(cache, "new");
        Check.Equal(2, calls, "Invalidation should request only the newest follow-up query.");
    }

    private static async Task ForgottenRefreshDoesNotPublish()
    {
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var finished = new ManualResetEventSlim();
        using var cache = new BackgroundRefreshCache<int, string>(key =>
        {
            started.Set();
            release.Wait();
            finished.Set();
            return "stale";
        });
        cache.Request(1);
        Check.That(started.Wait(2_000), "The refresh should start.");
        cache.Forget(1);
        release.Set();
        Check.That(finished.Wait(2_000), "The old query should finish.");
        await Task.Delay(30);
        Check.That(!cache.TryGet(1, out _), "A forgotten window must remain absent after its old query returns.");
    }

    private static async Task SlowRefreshDoesNotBlockAnotherWindow()
    {
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var finished = new ManualResetEventSlim();
        using var cache = new BackgroundRefreshCache<int, string>(key =>
        {
            if (key == 1)
            {
                started.Set();
                release.Wait();
                finished.Set();
                return "slow";
            }
            return "fast";
        });
        cache.Request(1);
        try
        {
            Check.That(started.Wait(2_000), "The slow window query should start.");
            cache.Request(2);
            await WaitForValue(cache, "fast", 2);
            Check.That(!cache.TryGet(1, out _), "The other window must finish while the slow query is still blocked.");
        }
        finally
        {
            release.Set();
            Check.That(finished.Wait(2_000), "The released query should finish.");
        }
    }

    private static Task RefreshConcurrencyIsBounded()
    {
        using var started = new CountdownEvent(2);
        using var finished = new CountdownEvent(2);
        using var excessQuery = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var calls = 0;
        using var cache = new BackgroundRefreshCache<int, string>(key =>
        {
            var current = Interlocked.Increment(ref calls);
            if (current > 2)
            {
                excessQuery.Set();
                return "excess";
            }
            started.Signal();
            release.Wait();
            finished.Signal();
            return "fresh";
        });
        try
        {
            for (var key = 1; key <= 128; key++)
                cache.Request(key);
            Check.That(started.Wait(2_000), "Different windows should be queried independently.");
            Check.That(!excessQuery.Wait(100), "A request storm must not create unbounded query workers.");
        }
        finally
        {
            cache.Dispose();
            release.Set();
            Check.That(SpinWait.SpinUntil(() => finished.CurrentCount == 2 - Volatile.Read(ref calls), 2_000),
                "All active queries should finish after disposal.");
        }
        return Task.CompletedTask;
    }

    private static async Task WaitForValue(BackgroundRefreshCache<int, string> cache, string expected, int key = 1)
    {
        var deadline = Environment.TickCount64 + 2_000;
        while (Environment.TickCount64 < deadline)
        {
            if (cache.TryGet(key, out var value) && value == expected)
                return;
            await Task.Delay(10);
        }
        throw new InvalidOperationException("The latest query result was not published.");
    }
}
