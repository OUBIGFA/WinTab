using System;
using System.Threading;
using WinTab.Helpers;

namespace WinTab.Hooks;

internal sealed class MergeOperation : IDisposable
{
    private readonly CancellationTokenSource _cancellation;
    private readonly Func<bool> _isOwnerCurrent;
    private readonly long _expiresAt;
    private int _disposed;

    public MergeOperation(WindowIdentity identity, int generation, CancellationToken lifetime,
        Func<bool> isOwnerCurrent, int timeoutMs, CancellationToken hookLifetime = default)
    {
        Identity = identity;
        Generation = generation;
        _isOwnerCurrent = isOwnerCurrent;
        _expiresAt = Environment.TickCount64 + Math.Max(1, timeoutMs);
        _cancellation = CancellationTokenSource.CreateLinkedTokenSource(lifetime, hookLifetime);
        Token = _cancellation.Token;
        _cancellation.CancelAfter(Math.Max(1, timeoutMs));
    }

    public WindowIdentity Identity { get; }
    public int Generation { get; }
    public CancellationToken Token { get; }
    public bool IsCurrent => Volatile.Read(ref _disposed) == 0 &&
        Environment.TickCount64 < _expiresAt && !Token.IsCancellationRequested && _isOwnerCurrent() && Identity.IsCurrent;

    public void ThrowIfInvalid()
    {
        if (!IsCurrent)
            throw new OperationCanceledException("The merge no longer owns this window.", Token);
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0)
            return;
        _cancellation.Cancel();
        _cancellation.Dispose();
    }
}
