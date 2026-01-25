using System;

namespace WinTab.Helpers;

internal sealed class BackoffDelay
{
    private readonly int _minMs;
    private readonly int _maxMs;
    private int _failures;

    public BackoffDelay(int minMs, int maxMs)
    {
        _minMs = Math.Max(1, minMs);
        _maxMs = Math.Max(_minMs, maxMs);
    }

    public int NextDelayMs()
    {
        var delay = _minMs * (1 << Math.Min(_failures, 6));
        if (delay > _maxMs)
            delay = _maxMs;

        _failures++;
        return delay;
    }

    public void Reset()
    {
        _failures = 0;
    }
}


