using System;
using System.Collections.Concurrent;

namespace WinTab.Helpers;

internal sealed class FailureTracker
{
    private readonly ConcurrentQueue<long> _failures = new();

    public void AddFailure()
    {
        _failures.Enqueue(StopwatchHelper.GetTimestamp());
        TrimOld();
    }

    public int CountRecent(int windowMs)
    {
        TrimOld(windowMs);
        return _failures.Count;
    }

    private void TrimOld(int windowMs = 5000)
    {
        while (_failures.TryPeek(out var ts) && StopwatchHelper.IsTimeUp(ts, windowMs))
            _failures.TryDequeue(out _);
    }
}


