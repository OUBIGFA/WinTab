using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

/// <summary>Pairs clicks and quarantines cancelled requests so late tabs cannot belong to the next click.</summary>
internal sealed class NavigationClickGate : IDisposable
{
    internal const int RequestLifetimeMs = 2_000;
    private readonly object _gate = new();
    private NavigationClickLease? _pending;
    private long _blockedUntil;
    private bool _disposed;

    public NavigationClickLease? Begin(Point point, long now)
    {
        lock (_gate)
        {
            if (_disposed)
                return null;
            if (_pending != null)
            {
                CancelCore();
                return null;
            }
            if (now < _blockedUntil)
                return null;
            return _pending = new NavigationClickLease(point, now + RequestLifetimeMs);
        }
    }

    public void Release(Point point, long now, int dragWidth, int dragHeight)
    {
        lock (_gate)
        {
            if (_pending is not { } pending)
                return;
            if (now >= pending.Deadline || Math.Abs((long)point.X - pending.Point.X) > dragWidth ||
                Math.Abs((long)point.Y - pending.Point.Y) > dragHeight)
            {
                CancelCore();
                return;
            }
            pending.Released.TrySetResult(true);
        }
    }

    public void Move(Point point, int dragWidth, int dragHeight)
    {
        lock (_gate)
        {
            if (_pending is { } pending && !pending.Released.Task.IsCompleted &&
                (Math.Abs((long)point.X - pending.Point.X) > dragWidth || Math.Abs((long)point.Y - pending.Point.Y) > dragHeight))
                CancelCore();
        }
    }

    public void Cancel()
    {
        lock (_gate)
            CancelCore();
    }

    private void CancelCore()
    {
        if (_pending is not { } pending)
            return;
        _blockedUntil = Math.Max(_blockedUntil, pending.Deadline);
        pending.Cancel();
    }

    public bool IsCurrent(NavigationClickLease lease, long now)
    {
        lock (_gate)
            return !_disposed && ReferenceEquals(_pending, lease) && !lease.Token.IsCancellationRequested && now < lease.Deadline;
    }

    public void Complete(NavigationClickLease lease, bool succeeded)
    {
        lock (_gate)
        {
            if (!ReferenceEquals(_pending, lease))
                return;
            if (!succeeded)
                _blockedUntil = Math.Max(_blockedUntil, lease.Deadline);
            _pending = null;
            lease.Dispose();
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            _disposed = true;
            CancelCore();
        }
    }
}

internal sealed class NavigationClickLease : IDisposable
{
    private readonly CancellationTokenSource _lifetime = new();
    public Point Point { get; }
    public long Deadline { get; }
    public CancellationToken Token { get; }
    public TaskCompletionSource<bool> Released { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

    public NavigationClickLease(Point point, long deadline)
    {
        Point = point;
        Deadline = deadline;
        Token = _lifetime.Token;
        _lifetime.CancelAfter(NavigationClickGate.RequestLifetimeMs);
    }

    public void Cancel() => _lifetime.Cancel();
    public void Dispose() => _lifetime.Dispose();
}
