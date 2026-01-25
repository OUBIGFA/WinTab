using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.Collections.Concurrent;

namespace WinTab.Helpers;

internal enum TelemetryLevel
{
    Debug,
    Info,
    Warn,
    Error
}

internal static class Telemetry
{
    private static readonly ConcurrentQueue<string> Queue = new();
    private static readonly ConcurrentDictionary<string, long> ThrottleTicks = new();
    private static readonly ConcurrentQueue<string> Recent = new();
    private static int _recentCount;
    private static readonly object FlushLock = new();
    private static readonly Timer FlushTimer;

    private const int MaxRecent = 200;

    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        Constants.AppName,
        "logs");

    private static readonly string LogFilePath = Path.Combine(LogDirectory, "app.log");

    static Telemetry()
    {
        FlushTimer = new Timer(_ => Flush(), null, 1000, 1000);
    }

    public static void Info(string message) => Write(TelemetryLevel.Info, message);
    public static void Warn(string message) => Write(TelemetryLevel.Warn, message);
    public static void Error(string message, Exception? ex = null)
    {
        var full = ex == null ? message : $"{message} | {ex.GetType().Name}: {ex.Message}";
        Write(TelemetryLevel.Error, full);
    }

    public static void ThrottledInfo(string key, string message, int windowMs = 1000)
    {
        if (IsThrottled(key, windowMs)) return;
        Write(TelemetryLevel.Info, message);
    }

    public static void ThrottledWarn(string key, string message, int windowMs = 1000)
    {
        if (IsThrottled(key, windowMs)) return;
        Write(TelemetryLevel.Warn, message);
    }

    private static void Write(TelemetryLevel level, string message)
    {
        try
        {
            var timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            var line = $"{timestamp} [{level}] {message}";
            Queue.Enqueue(line);
            AddRecent(line);
        }
        catch
        {
            // Never throw from telemetry
        }
    }

    private static bool IsThrottled(string key, int windowMs)
    {
        var now = Stopwatch.GetTimestamp();
        if (ThrottleTicks.TryGetValue(key, out var last) && !StopwatchHelper.IsTimeUp(last, windowMs))
            return true;

        ThrottleTicks[key] = now;
        return false;
    }

    public static string[] GetRecent(int max = 100)
    {
        var items = Recent.ToArray();
        if (items.Length <= max) return items;

        var start = Math.Max(0, items.Length - max);
        return items.Skip(start).ToArray();
    }

    private static void AddRecent(string line)
    {
        Recent.Enqueue(line);
        var count = Interlocked.Increment(ref _recentCount);
        while (count > MaxRecent && Recent.TryDequeue(out _))
        {
            count = Interlocked.Decrement(ref _recentCount);
        }
    }

    private static void Flush()
    {
        if (Queue.IsEmpty) return;

        lock (FlushLock)
        {
            if (Queue.IsEmpty) return;

            try
            {
                Directory.CreateDirectory(LogDirectory);

                var sb = new StringBuilder();
                while (Queue.TryDequeue(out var line))
                    sb.AppendLine(line);

                File.AppendAllText(LogFilePath, sb.ToString());
            }
            catch
            {
                // Never throw from telemetry
            }
        }
    }
}


