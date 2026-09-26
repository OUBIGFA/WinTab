using System;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

/// <summary>
/// Pairs a middle click's button-down with its button-up and follows one click at a time. A new click retires
/// an unresolved older one instead of being dropped: a click that opened nothing, or one that was cancelled,
/// must never cost the next click its activation. Which tab belongs to a click is decided by the tabs Explorer
/// adds, not by the pointer: Explorer opens the tab even when the pointer or the wheel moves during the click.
/// Explorer opens the tab after the button is released, and a busy Explorer or a drive that has to wake up can
/// take seconds to do so; the click therefore waits for its tab from the release, not from the button-down.
/// </summary>
internal sealed class NavigationClickGate : IDisposable
{
    /// <summary>How long the button may stay down before the click is given up.</summary>
    internal const int HoldLimitMs = 10_000;
    /// <summary>How long after the release the click waits for Explorer to open its tab.</summary>
    internal const int AfterReleaseMs = 8_000;
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
            return _pending = new NavigationClickLease(window, now);
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
            pending.MarkReleased(now);
        }
    }

    public bool Cancel()
    {
        lock (_gate)
            return CancelCore();
    }

    /// <summary>
    /// The pending lease owns its window; retiring an older click cannot change that ownership. A window the
    /// click's window owns, such as one of Explorer's popups, is still that window, and no foreground window at
    /// all is a transition, not a switch.
    /// </summary>
    public bool CancelIfForegroundChanged(nint window, nint owner = 0)
    {
        lock (_gate)
            return window != 0 && _pending is { } pending && pending.Window != window && pending.Window != owner && CancelCore();
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
    private long _deadline;
    public nint Window { get; }
    /// <summary>The click ends at this tick: <see cref="NavigationClickGate.HoldLimitMs"/> after the button-down, then <see cref="NavigationClickGate.AfterReleaseMs"/> after the release.</summary>
    public long Deadline => Interlocked.Read(ref _deadline);
    public CancellationToken Token { get; }
    public TaskCompletionSource<bool> Released { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

    public NavigationClickLease(nint window, long now)
    {
        Window = window;
        _deadline = now + NavigationClickGate.HoldLimitMs;
        Token = _lifetime.Token;
        _lifetime.CancelAfter(NavigationClickGate.HoldLimitMs);
    }

    /// <summary>Called under the gate's lock, so it cannot race the lease's cancellation or disposal.</summary>
    internal void MarkReleased(long now)
    {
        Interlocked.Exchange(ref _deadline, now + NavigationClickGate.AfterReleaseMs);
        if (!_lifetime.IsCancellationRequested)
            _lifetime.CancelAfter(NavigationClickGate.AfterReleaseMs);
        Released.TrySetResult(true);
    }

    public void Cancel() => _lifetime.Cancel();
    public void Dispose() => _lifetime.Dispose();
}
