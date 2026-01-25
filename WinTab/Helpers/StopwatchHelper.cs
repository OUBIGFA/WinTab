using System;
using System.Diagnostics;

namespace WinTab.Helpers;

internal static class StopwatchHelper
{
    public static long GetTimestamp() => Stopwatch.GetTimestamp();

    public static bool IsTimeUp(long startTicks, int timeMs)
    {
#if NET7_0_OR_GREATER
        var elapsedTime = Stopwatch.GetElapsedTime(startTicks);
#else
        var tickFrequency = (double)10_000_000 / Stopwatch.Frequency;
        var elapsedTime = new TimeSpan((long)((Stopwatch.GetTimestamp() - startTicks) * tickFrequency));
#endif
        return elapsedTime.TotalMilliseconds >= timeMs;
    }
}


