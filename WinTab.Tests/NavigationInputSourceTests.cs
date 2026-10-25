using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using WinTab.Hooks;

internal static class NavigationInputSourceTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("navigation input source shares one observer across independent subscriptions", SharesObserver);
        yield return ("navigation input source isolates subscriber failures", IsolatesFailures);
        yield return ("navigation input source disposes once and rejects later subscriptions", DisposesOnce);
    }

    private sealed class FakeObserver : IDisposable
    {
        private readonly Action<NavigationPointerInput> _pointer;
        private readonly Action _cancel;
        private readonly Action<nint> _foreground;
        internal int DisposeCalls;

        internal FakeObserver(Action<NavigationPointerInput> pointer, Action cancel, Action<nint> foreground)
        {
            _pointer = pointer;
            _cancel = cancel;
            _foreground = foreground;
        }

        internal void Emit()
        {
            _pointer(new NavigationPointerInput(NavigationPointerKind.LeftDown, new Point(1, 2), false));
            _cancel();
            _foreground(123);
        }

        public void Dispose() => DisposeCalls++;
    }

    private static Task SharesObserver()
    {
        var creates = 0;
        var observer = default(FakeObserver);
        using var source = new NavigationInputSource((pointer, cancel, foreground) =>
        {
            creates++;
            return observer = new FakeObserver(pointer, cancel, foreground);
        });
        var firstCalls = 0;
        var secondCalls = 0;
        var first = source.Subscribe(_ => firstCalls++, () => { }, _ => { });
        var second = source.Subscribe(_ => secondCalls++, () => { }, _ => { });
        Check.Equal(1, creates, "The first subscriber creates the only observer");
        observer!.Emit();
        Check.Equal(1, firstCalls, "The first subscriber receives input");
        Check.Equal(1, secondCalls, "The second subscriber receives the same input");
        first.Dispose();
        Check.Equal(0, observer.DisposeCalls, "The observer survives while another feature remains");
        observer.Emit();
        Check.Equal(1, firstCalls, "A removed subscriber receives nothing");
        Check.Equal(2, secondCalls, "The remaining subscriber continues receiving input");
        second.Dispose();
        Check.Equal(1, observer.DisposeCalls, "The last subscription releases the observer");
        return Task.CompletedTask;
    }

    private static Task IsolatesFailures()
    {
        var observer = default(FakeObserver);
        using var source = new NavigationInputSource((pointer, cancel, foreground) =>
            observer = new FakeObserver(pointer, cancel, foreground));
        var calls = 0;
        using var failing = source.Subscribe(_ => throw new InvalidOperationException("test"), () => { }, _ => { });
        using var healthy = source.Subscribe(_ => calls++, () => { }, _ => { });
        observer!.Emit();
        Check.Equal(1, calls, "One subscriber's exception must not block the other subscriber");
        return Task.CompletedTask;
    }

    private static Task DisposesOnce()
    {
        var observer = default(FakeObserver);
        var source = new NavigationInputSource((pointer, cancel, foreground) =>
            observer = new FakeObserver(pointer, cancel, foreground));
        var subscription = source.Subscribe(_ => { }, () => { }, _ => { });
        source.Dispose();
        source.Dispose();
        subscription.Dispose();
        Check.Equal(1, observer!.DisposeCalls, "Disposal is idempotent across the source and subscriptions");
        Check.Throws<ObjectDisposedException>(() => source.Subscribe(_ => { }, () => { }, _ => { }),
            "A disposed source cannot create a new observer");
        return Task.CompletedTask;
    }
}
