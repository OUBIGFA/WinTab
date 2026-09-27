using System;
using System.Threading.Tasks;

namespace WinTab.Hooks;

/// <summary>
/// Re-conceals merge sources for a short while after Explorer touched a window, because Explorer can show a
/// concealed frame again at any time during a merge. Every start is followed by at least one scan and asks
/// for its duration; a stream of starts cannot keep one pulse running past its absolute ceiling. Whether the
/// pulse continues and whether a start extends it are decided under one lock, so a start that arrives while
/// the pulse is ending either extends it or starts the next one; it is never lost.
/// </summary>
internal sealed class MergeSourceConcealPulse
{
    private const int DefaultAbsoluteCeilingMs = 1_500;
    private const int DefaultSleepMs = 25;

    private readonly object _gate = new();
    private readonly int _absoluteCeilingMs;
    private readonly int _sleepMs;
    private bool _running;
    /// <summary>A start has not been followed by a scan yet.</summary>
    private bool _scanPending;
    private long _startedAt;
    private long _until;

    public MergeSourceConcealPulse()
        : this(DefaultAbsoluteCeilingMs, DefaultSleepMs)
    {
    }

    internal MergeSourceConcealPulse(int absoluteCeilingMs, int sleepMs)
    {
        _absoluteCeilingMs = Math.Max(1, absoluteCeilingMs);
        _sleepMs = Math.Max(1, sleepMs);
    }

    public void Start(Func<bool> isEnabled, Action concealOnce, int durationMs = 1_200)
    {
        lock (_gate)
        {
            if (!isEnabled())
                return;

            var now = Environment.TickCount64;
            if (!_running)
                _startedAt = now;
            _until = Math.Max(_until, Math.Min(now + Math.Max(1, durationMs), _startedAt + _absoluteCeilingMs));
            _scanPending = true;
            if (_running)
                return;
            _running = true;
        }

        _ = Task.Run(() => RunAsync(isEnabled, concealOnce));
    }

    private async Task RunAsync(Func<bool> isEnabled, Action concealOnce)
    {
        while (true)
        {
            lock (_gate)
            {
                if (!isEnabled() || (!_scanPending && Environment.TickCount64 >= _until))
                {
                    _running = false;
                    _until = 0;
                    return;
                }
                _scanPending = false;
            }

            try
            {
                concealOnce();
            }
            catch
            {
                // Keep the pulse bounded even if one scan fails.
            }

            await Task.Delay(_sleepMs);
        }
    }

    public void Stop()
    {
        lock (_gate)
        {
            _until = 0;
            _scanPending = false;
        }
    }
}
