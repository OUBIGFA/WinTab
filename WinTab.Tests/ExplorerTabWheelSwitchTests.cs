using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Hooks;
using WinTab.WinAPI;

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
        yield return ("cold wheel hit testing retains the first gesture", ColdHitRetainsFirstGesture);
        yield return ("file-view wheel input does not switch or consume the throttle", FileViewDoesNotConsumeThrottle);
        yield return ("queued wheel input is throttled by arrival time", QueuedInputUsesArrivalTime);
        yield return ("slow hit testing does not replay a rapid wheel notch", SlowHitDoesNotReplayRapidInput);
        yield return ("stopping and restarting wheel switching discards the old generation", RestartDiscardsPendingInput);
        yield return ("invalid wheel gestures never switch after hit testing", InvalidInputDoesNotSwitch);
        yield return ("expired wheel gestures never switch after hit testing", ExpiredInputDoesNotSwitch);
        yield return ("wheel hit-test failure does not poison the next input", HitTestFailureDoesNotPoisonQueue);
        yield return ("wheel switch failure does not poison the next input", SwitchFailureDoesNotPoisonQueue);
        yield return ("repeated cold tab-row queries preserve the in-flight read", RepeatedColdQueriesShareRead);
    }

    private static readonly Rectangle TabRow = new(10, 20, 600, 30);
    private static readonly Point TabPoint = new(40, 30);
    private static readonly TimeSpan TestTimeout = TimeSpan.FromSeconds(3);

    private static TaskCompletionSource<T> Completion<T>() => new(TaskCreationOptions.RunContinuationsAsynchronously);

    private static async Task ColdHitRetainsFirstGesture()
    {
        var row = Completion<Rectangle?>();
        var reading = Completion<bool>();
        using var fixture = new WheelFixture(_ =>
        {
            reading.TrySetResult(true);
            return row.Task;
        });

        var first = fixture.Queue();
        await reading.Task.WaitAsync(TestTimeout);
        Check.That(!first.IsCompleted, "The first wheel input must wait for the cold hit test.");
        Check.Equal(0, fixture.Steps.Count, "A switch cannot precede the hit test.");
        row.SetResult(TabRow);

        Check.That(await first, "The original input must switch without a second wheel event.");
        Check.Equal(1, fixture.ReadCount);
        Check.Equal(1, fixture.Steps.Count);
        Check.Equal(1, fixture.Steps[0], "Wheel down must select the next tab.");
    }

    private static async Task FileViewDoesNotConsumeThrottle()
    {
        var row = Completion<Rectangle?>();
        var reading = Completion<bool>();
        using var fixture = new WheelFixture(_ =>
        {
            reading.TrySetResult(true);
            return row.Task;
        });
        var arrivedAt = Environment.TickCount64 - 100;
        var fileScroll = fixture.Queue(arrivedAt, point: new Point(TabPoint.X, TabRow.Bottom + 100));
        await reading.Task.WaitAsync(TestTimeout);
        var tabScroll = fixture.Queue(arrivedAt + 10);
        row.SetResult(TabRow);

        Check.That(!await fileScroll, "A cold hit in the file area must not switch tabs.");
        Check.That(await tabScroll, "File scrolling must not hold off the immediately following tab-row notch.");
        Check.Equal(1, fixture.Steps.Count, "Only the tab-row gesture may issue a switch.");
    }

    private static async Task QueuedInputUsesArrivalTime()
    {
        var row = Completion<Rectangle?>();
        var reading = Completion<bool>();
        using var fixture = new WheelFixture(_ =>
        {
            reading.TrySetResult(true);
            return row.Task;
        });
        // The inputs arrived at different times, but all their hit tests finish together.
        var arrivedAt = Environment.TickCount64 - 400;
        var first = fixture.Queue(arrivedAt);
        await reading.Task.WaitAsync(TestTimeout);
        var rapid = fixture.Queue(arrivedAt + 50);
        var afterHold = fixture.Queue(arrivedAt + 250, delta: 120);
        row.SetResult(TabRow);

        Check.That(await first, "The first queued notch must switch.");
        Check.That(!await rapid, "A notch that arrived inside the hold must be dropped.");
        Check.That(await afterHold, "A notch outside the arrival-time hold must survive compressed UIA completion.");
        Check.Equal(2, fixture.Steps.Count);
        Check.Equal(1, fixture.Steps[0]);
        Check.Equal(-1, fixture.Steps[1], "Queued gestures retain their arrival order and direction.");
    }

    private static async Task SlowHitDoesNotReplayRapidInput()
    {
        var row = Completion<Rectangle?>();
        var reading = Completion<bool>();
        using var fixture = new WheelFixture();
        var arrivedAt = Environment.TickCount64 - 100;
        Check.That(await fixture.Queue(arrivedAt), "The first notch starts the hold.");
        fixture.ReadRow = _ =>
        {
            reading.TrySetResult(true);
            return row.Task;
        };
        var rapid = fixture.Queue(arrivedAt + 50);
        await reading.Task.WaitAsync(TestTimeout);
        // The second UIA result takes longer than the medium hold, but the input arrived inside it.
        await Task.Delay(250);
        row.SetResult(TabRow);

        Check.That(!await rapid, "Slow UIA must not turn a fast flick into a delayed extra switch.");
        Check.Equal(1, fixture.Steps.Count);
    }

    private static async Task RestartDiscardsPendingInput()
    {
        var row = Completion<Rectangle?>();
        var reading = Completion<CancellationToken>();
        using var fixture = new WheelFixture();
        var arrivedAt = Environment.TickCount64 - 100;
        Check.That(await fixture.Queue(arrivedAt), "A pre-stop switch establishes the throttle state.");
        fixture.ReadRow = token =>
        {
            reading.TrySetResult(token);
            // Model UIA that finishes even after cancellation was requested.
            return row.Task;
        };
        var inFlight = fixture.Queue(arrivedAt + 10);
        var queued = fixture.Queue(arrivedAt + 20);
        var oldToken = await reading.Task.WaitAsync(TestTimeout);
        fixture.Controller.Stop();
        Check.That(oldToken.IsCancellationRequested, "Stopping must cancel the pending hit test's lifetime.");
        Check.That(!await fixture.Queue(), "Input received while stopped must be rejected.");
        fixture.Controller.Start();
        var restarted = fixture.Queue(arrivedAt + 30);
        row.SetResult(TabRow);

        Check.That(!await inFlight && !await queued, "Restart must not revive the old generation's gestures.");
        Check.That(await restarted, "The restarted controller must process new input with a reset throttle.");
        Check.Equal(3, fixture.ReadCount, "Stopped and obsolete queued input must not start another hit test.");
        Check.Equal(2, fixture.Steps.Count, "Only the pre-stop and new-generation gestures may switch.");
    }

    private static async Task InvalidInputDoesNotSwitch()
    {
        var row = Completion<Rectangle?>();
        var reading = Completion<bool>();
        using var fixture = new WheelFixture(_ =>
        {
            reading.TrySetResult(true);
            return row.Task;
        });
        Check.That(!await fixture.Queue(isCurrent: () => false), "An already-invalid target must be rejected.");
        Check.Equal(0, fixture.ReadCount, "Invalid input must not start UIA.");
        var current = 1;
        var arrivedAt = Environment.TickCount64;
        var inFlight = fixture.Queue(arrivedAt, isCurrent: () => Volatile.Read(ref current) != 0);
        var queued = fixture.Queue(arrivedAt, isCurrent: () => Volatile.Read(ref current) != 0);
        await reading.Task.WaitAsync(TestTimeout);
        Volatile.Write(ref current, 0);
        row.SetResult(TabRow);

        Check.That(!await inFlight && !await queued, "A moved, closed or replaced target must not switch later.");
        Check.Equal(0, fixture.Steps.Count);
        Check.That(await fixture.Queue(arrivedAt), "Rejected gestures must not consume the next valid input's throttle.");
        Check.Equal(2, fixture.ReadCount, "The invalid queued gesture must be checked before another read.");
    }

    private static async Task ExpiredInputDoesNotSwitch()
    {
        var row = Completion<Rectangle?>();
        var reading = Completion<bool>();
        using var fixture = new WheelFixture(_ =>
        {
            reading.TrySetResult(true);
            return row.Task;
        });
        Check.That(!await fixture.Queue(Environment.TickCount64 - 2_000), "An already-expired gesture must be rejected.");
        Check.Equal(0, fixture.ReadCount);
        var arrivedAt = Environment.TickCount64;
        var inFlight = fixture.Queue(arrivedAt);
        var queued = fixture.Queue(arrivedAt);
        await reading.Task.WaitAsync(TestTimeout);
        // Exercise expiry after awaiting UIA, not just rejection at enqueue time.
        await Task.Delay(1_050);
        row.SetResult(TabRow);

        Check.That(!await inFlight && !await queued, "Expired gestures must not replay when a slow read finally completes.");
        Check.Equal(0, fixture.Steps.Count);
        Check.That(await fixture.Queue(), "A fresh gesture must still work after expired work drains.");
        Check.Equal(2, fixture.ReadCount, "Expired queued input must not perform an additional hit test.");
    }

    private static async Task HitTestFailureDoesNotPoisonQueue()
    {
        var row = Completion<Rectangle?>();
        var reading = Completion<bool>();
        var failure = new InvalidOperationException("Injected tab-row read failure.");
        using var fixture = new WheelFixture(_ =>
        {
            reading.TrySetResult(true);
            return row.Task;
        });
        var arrivedAt = Environment.TickCount64;
        var failed = fixture.Queue(arrivedAt);
        await reading.Task.WaitAsync(TestTimeout);
        var next = fixture.Queue(arrivedAt);
        fixture.ReadRow = _ => Task.FromResult<Rectangle?>(TabRow);
        row.SetException(failure);

        await ExpectFailure(failed, failure);
        Check.That(await next, "A faulted queue tail must not block or throttle the next input.");
        Check.Equal(1, fixture.Steps.Count);
    }

    private static async Task SwitchFailureDoesNotPoisonQueue()
    {
        foreach (var throws in new[] { false, true })
        {
            using var fixture = new WheelFixture();
            var failure = new InvalidOperationException("Injected switch failure.");
            fixture.SwitchTab = () => throws ? Task.FromException<bool>(failure) : Task.FromResult(false);
            var arrivedAt = Environment.TickCount64 - 400;
            var failed = fixture.Queue(arrivedAt);
            if (throws)
                await ExpectFailure(failed, failure);
            else
                Check.That(!await failed, "A failed switch must be reported as unsuccessful.");

            fixture.SwitchTab = () => Task.FromResult(true);
            Check.That(await fixture.Queue(arrivedAt + 250), "A later notch must work after a false result or exception.");
            Check.Equal(2, fixture.Steps.Count);
        }
    }

    private static async Task ExpectFailure(Task<bool> work, Exception expected)
    {
        try
        {
            await work;
        }
        catch (Exception actual) when (ReferenceEquals(actual, expected))
        {
            return;
        }
        throw new InvalidOperationException("The injected failure must reach the caller, not become a successful result.");
    }

    private static async Task RepeatedColdQueriesShareRead()
    {
        // Cabinet-class test frames live on their own desktop, never the user's input desktop.
        using var frame = new RemoteExplorerFrame(visible: false, explorerClass: true);
        Check.That(WinApi.GetWindowRect(frame.Handle, out var bounds), "The isolated frame must have bounds.");
        var expected = Rectangle.FromLTRB(bounds.Left, bounds.Top, bounds.Right, bounds.Top + 24);
        var reading = Completion<bool>();
        var release = Completion<bool>();
        var reads = 0;
        using var tester = new TabStripHitTester(_ =>
        {
            Interlocked.Increment(ref reads);
            reading.TrySetResult(true);
            release.Task.WaitAsync(TestTimeout).GetAwaiter().GetResult();
            return [new ExplorerTabAutomation.Tab(new System.Windows.Rect(bounds.Left, bounds.Top, 80, 24), true)];
        });
        try
        {
            Check.That(!tester.TryGetTabRow(frame.Handle, out _), "The first lookup must start a cold read.");
            await reading.Task.WaitAsync(TestTimeout);
            for (var index = 0; index < 8; index++)
                Check.That(!tester.TryGetTabRow(frame.Handle, out _), "Repeated lookups remain cold until the reader returns.");
            var pending = tester.GetTabRowAsync(frame.Handle, CancellationToken.None);
            release.SetResult(true);

            Check.Equal<Rectangle?>(expected, await pending.WaitAsync(TestTimeout), "The original computation must populate the row.");
            Check.Equal(1, Volatile.Read(ref reads), "Cold polling must not invalidate and recompute the in-flight result.");
            Check.That(tester.TryGetTabRow(frame.Handle, out var cached), "The published row must be immediately reusable.");
            Check.Equal(expected, cached);
        }
        finally
        {
            release.TrySetResult(true);
        }
    }

    private sealed class WheelFixture : IDisposable
    {
        private static readonly nint Window = (nint)42;
        public ExplorerTabWheelSwitchController Controller { get; }
        public Func<CancellationToken, Task<Rectangle?>> ReadRow { get; set; }
        public Func<Task<bool>> SwitchTab { get; set; } = () => Task.FromResult(true);
        public List<int> Steps { get; } = [];
        public int ReadCount { get; private set; }

        public WheelFixture(Func<CancellationToken, Task<Rectangle?>>? readRow = null)
        {
            ReadRow = readRow ?? (_ => Task.FromResult<Rectangle?>(TabRow));
            Controller = new ExplorerTabWheelSwitchController((window, token) =>
            {
                Check.Equal(Window, window, "The hit test must target the input's window.");
                ReadCount++;
                return ReadRow(token).WaitAsync(TestTimeout);
            }, (window, step, isCurrent) =>
            {
                Check.Equal(Window, window, "The switch must target the input's window.");
                Check.That(isCurrent(), "The switch callback must receive a current gesture.");
                Steps.Add(step);
                return SwitchTab();
            });
        }

        public Task<bool> Queue(long? receivedAt = null, Point? point = null, int delta = -120, Func<bool>? isCurrent = null) =>
            Controller.QueueAsync(Window, point ?? TabPoint, delta, receivedAt ?? Environment.TickCount64,
                WheelSwitchSensitivity.Medium, isCurrent ?? (() => true)).WaitAsync(TestTimeout);

        public void Dispose() => Controller.Dispose();
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
