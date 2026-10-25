using System;
using System.Collections.Generic;
using System.Threading;

namespace WinTab.Hooks;

/// <summary>
/// Shares the process-wide Raw Input registration between navigation features. A subscription does not
/// own the other feature's lifetime. Based on hbiui/WinTab 4121b08, with allocation-free input dispatch.
/// </summary>
internal sealed class NavigationInputSource : IDisposable
{
    private readonly object _gate = new();
    private readonly List<Subscription> _subscriptions = [];
    private readonly Func<Action<NavigationPointerInput>, Action, Action<nint>, IDisposable> _createObserver;
    private Subscription[] _snapshot = [];
    private IDisposable? _observer;
    private bool _disposed;

    public NavigationInputSource()
        : this((pointer, cancel, foreground) => new NavigationInputObserver(pointer, cancel, foreground)) { }

    internal NavigationInputSource(Func<Action<NavigationPointerInput>, Action, Action<nint>, IDisposable> createObserver) =>
        _createObserver = createObserver;

    internal IDisposable Subscribe(Action<NavigationPointerInput> pointer, Action cancel, Action<nint> foreground)
    {
        var subscription = new Subscription(this, pointer, cancel, foreground);
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _observer ??= _createObserver(DispatchPointer, DispatchCancel, DispatchForeground);
            _subscriptions.Add(subscription);
            Volatile.Write(ref _snapshot, _subscriptions.ToArray());
        }
        return subscription;
    }

    private void DispatchPointer(NavigationPointerInput input)
    {
        foreach (var subscription in Volatile.Read(ref _snapshot))
        {
            if (!subscription.Active) continue;
            try { subscription.Pointer(input); }
            catch (Exception exception) { ExplorerDebugLog.Write($"Navigation pointer subscriber failed: {exception.Message}"); }
        }
    }

    private void DispatchCancel()
    {
        foreach (var subscription in Volatile.Read(ref _snapshot))
        {
            if (!subscription.Active) continue;
            try { subscription.Cancel(); }
            catch (Exception exception) { ExplorerDebugLog.Write($"Navigation cancellation subscriber failed: {exception.Message}"); }
        }
    }

    private void DispatchForeground(nint window)
    {
        foreach (var subscription in Volatile.Read(ref _snapshot))
        {
            if (!subscription.Active) continue;
            try { subscription.Foreground(window); }
            catch (Exception exception) { ExplorerDebugLog.Write($"Navigation foreground subscriber failed: {exception.Message}"); }
        }
    }

    private void Unsubscribe(Subscription subscription)
    {
        IDisposable? observer = null;
        lock (_gate)
        {
            if (!subscription.Active) return;
            subscription.Active = false;
            _subscriptions.Remove(subscription);
            Volatile.Write(ref _snapshot, _subscriptions.ToArray());
            if (_subscriptions.Count == 0)
            {
                observer = _observer;
                _observer = null;
            }
        }
        // Never join the input thread while holding a lock its subscriber may need.
        observer?.Dispose();
    }

    public void Dispose()
    {
        IDisposable? observer;
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true;
            foreach (var subscription in _subscriptions) subscription.Active = false;
            _subscriptions.Clear();
            Volatile.Write(ref _snapshot, []);
            observer = _observer;
            _observer = null;
        }
        observer?.Dispose();
    }

    private sealed class Subscription(NavigationInputSource owner, Action<NavigationPointerInput> pointer,
        Action cancel, Action<nint> foreground) : IDisposable
    {
        internal volatile bool Active = true;
        internal Action<NavigationPointerInput> Pointer { get; } = pointer;
        internal Action Cancel { get; } = cancel;
        internal Action<nint> Foreground { get; } = foreground;
        public void Dispose() => owner.Unsubscribe(this);
    }
}
