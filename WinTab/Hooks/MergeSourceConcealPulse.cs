using System;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Hooks;

internal sealed class MergeSourceConcealPulse
{
    private const int DefaultAbsoluteCeilingMs = 1_500;
    private const int DefaultSleepMs = 25;

    private readonly int _absoluteCeilingMs;
    private readonly int _sleepMs;
    private int _running;
    private long _untilTicks;
    private long _firstStartTicks;

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
        if (!isEnabled())
            return;

        var requestedUntilTicks = DateTime.UtcNow.AddMilliseconds(Math.Max(1, durationMs)).Ticks;
        Interlocked.CompareExchange(ref _firstStartTicks, DateTime.UtcNow.Ticks, 0);
        var firstStart = Volatile.Read(ref _firstStartTicks);
        var ceilingTicks = firstStart + _absoluteCeilingMs * TimeSpan.TicksPerMillisecond;
        var untilTicks = Math.Min(requestedUntilTicks, ceilingTicks);

        var currentTicks = Volatile.Read(ref _untilTicks);
        while (untilTicks > currentTicks)
        {
            var previousTicks = Interlocked.CompareExchange(ref _untilTicks, untilTicks, currentTicks);
            if (previousTicks == currentTicks)
                break;

            currentTicks = previousTicks;
        }

        if (Interlocked.Exchange(ref _running, 1) == 1)
            return;

        _ = Task.Run(async () =>
        {
            try
            {
                while (DateTime.UtcNow.Ticks < Volatile.Read(ref _untilTicks))
                {
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
            finally
            {
                Interlocked.Exchange(ref _running, 0);
                Interlocked.Exchange(ref _firstStartTicks, 0);
                Interlocked.Exchange(ref _untilTicks, 0);
            }
        });
    }

    public void Stop()
    {
        Interlocked.Exchange(ref _untilTicks, 0);
    }
}
