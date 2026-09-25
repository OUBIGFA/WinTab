using System;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

/// <summary>
/// Pairs a middle click's button-down with its button-up and follows one click at a time. A new click retires
/// an unresolved older one instead of being dropped: a click that opened nothing, or one that was cancelled,
/// must never cost the next click its activation. Which tab belongs to a click is decided by the tabs Explorer
/// adds, not by the pointer: Explorer opens the tab even when the pointer or the wheel moves during the click.
/// </summary>
internal sealed class NavigationClickGate : IDisposable
{
    internal const int RequestLifetimeMs = 2_000;
    private readonly object _gate = new();
    private NavigationClickLease? _pending;
    private bool _disposed;

    /// <summary>Starts following a click. <paramref name="replaced"/> tells whether an unresolved click was retired for it.</summary>
    public NavigationClickLease? Begin(nint window, long now, out bool replaced)
    {
        lock (_gate)
        {
            replaced = false;
            if (_disposed)
                return null;
            if (_pending is { } older)
            {
                replaced = now < older.Deadline;
                older.Cancel();
            }
            return _pending = new NavigationClickLease(window, now + RequestLifetimeMs);
        }
    }

    public void Release(long now)
    {
        lock (_gate)
        {
            if (_pending is not { } pending)
                return;
            if (now >= pending.Deadline)
            {
                CancelCore();
                return;
            }
            pending.Released.TrySetResult(true);
        }
    }

    public bool Cancel()
    {
        lock (_gate)
            return CancelCore();
    }

    /// <summary>The pending lease owns its window; retiring an older click cannot change that ownership.</summary>
    public bool CancelIfForegroundChanged(nint window)
    {
        lock (_gate)
            return _pending is { } pending && pending.Window != window && CancelCore();
    }

    private bool CancelCore()
    {
        if (_pending is not { } pending)
            return false;
        _pending = null;
        pending.Cancel();
        return true;
    }

    public bool IsCurrent(NavigationClickLease lease, long now)
    {
        lock (_gate)
            return !_disposed && ReferenceEquals(_pending, lease) && !lease.Token.IsCancellationRequested && now < lease.Deadline;
    }

    /// <summary>The click's worker is done with it. A click that was retired meanwhile leaves the current one alone.</summary>
    public void Complete(NavigationClickLease lease)
    {
        lock (_gate)
        {
            if (ReferenceEquals(_pending, lease))
                _pending = null;
        }
        lease.Dispose();
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
    public nint Window { get; }
    public long Deadline { get; }
    public CancellationToken Token { get; }
    public TaskCompletionSource<bool> Released { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

    public NavigationClickLease(nint window, long deadline)
    {
        Window = window;
        Deadline = deadline;
        Token = _lifetime.Token;
        _lifetime.CancelAfter(NavigationClickGate.RequestLifetimeMs);
    }

    public void Cancel() => _lifetime.Cancel();
    public void Dispose() => _lifetime.Dispose();
}
