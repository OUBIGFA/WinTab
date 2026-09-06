using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using WinTab.WinAPI;

namespace WinTab.Helpers;

public static class Helper
{
    private static int _lastCtrlShiftCheckAt;
    private static bool _lastCtrlShiftCheckValue;

    public static Task<T> DoUntilNotDefaultAsync<T>(Func<Task<T>> action, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        return DoUntilConditionAsync(
            action,
            result => !EqualityComparer<T?>.Default.Equals(result, default),
            timeMs,
            sleepMs,
            cancellationToken);
    }
    public static Task<T> DoUntilNotDefaultAsync<T>(Func<T> action, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        return DoUntilConditionAsync(
            action,
            result => !EqualityComparer<T?>.Default.Equals(result, default),
            timeMs,
            sleepMs,
            cancellationToken);
    }
    public static Task<T> DoUntilConditionAsync<T>(Func<T> action, Predicate<T> predicate, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default) =>
        DoUntilConditionAsync(() => Task.FromResult(action()), predicate, timeMs, sleepMs, cancellationToken);
    public static async Task<T> DoUntilConditionAsync<T>(Func<Task<T>> action, Predicate<T> predicate, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        var startTicks = Stopwatch.GetTimestamp();
        var remainingMs = Math.Max(1, timeMs);
        T result;
        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            result = await action().WaitAsync(TimeSpan.FromMilliseconds(remainingMs), cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (predicate(result))
                return result;

            remainingMs = timeMs - (int)Stopwatch.GetElapsedTime(startTicks).TotalMilliseconds;
            if (remainingMs <= 0)
                break;
            await Task.Delay(Math.Min(Math.Max(1, sleepMs), remainingMs), cancellationToken).ConfigureAwait(false);
            remainingMs = timeMs - (int)Stopwatch.GetElapsedTime(startTicks).TotalMilliseconds;
        }
        while (remainingMs > 0);

        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    public static bool IsTimeUp(long startTicks, int timeMs)
    {
        return Stopwatch.GetElapsedTime(startTicks).TotalMilliseconds >= timeMs;
    }

    public static Icon? GetIcon() => Icon.ExtractAssociatedIcon(GetExecutablePath());

    public static bool IsCtrlShiftDown()
    {
        if (_lastCtrlShiftCheckValue && Environment.TickCount - _lastCtrlShiftCheckAt < 1_000)
            return true;

        _lastCtrlShiftCheckValue =
            (KeyboardSimulator.IsKeyPressed((int)VirtualKey.LeftControl) || KeyboardSimulator.IsKeyPressed((int)VirtualKey.RightControl)) &&
               (KeyboardSimulator.IsKeyPressed((int)VirtualKey.LeftShift) || KeyboardSimulator.IsKeyPressed((int)VirtualKey.RightShift));

        _lastCtrlShiftCheckAt = Environment.TickCount;
        return _lastCtrlShiftCheckValue;
    }
    public static void BypassWinForegroundRestrictions()
    {
        // Simulate a key press to bypass the Foreground restriction
        // https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow#remarks
        KeyboardSimulator.SendKeyPress(VirtualKey.F23);
    }

    public static void RestoreWindowToForeground(nint window)
    {
        if (WinApi.IsIconic(window))
            WinApi.ShowWindow(window, WinApi.SW_SHOWNOACTIVATE);

        if (WinApi.SetForegroundWindow(window))
            return;

        BypassWinForegroundRestrictions();
        WinApi.SetForegroundWindow(window);
    }

    public static string NormalizeLocation(string location)
    {
        if (location.IndexOf('%') > -1)
            location = Environment.ExpandEnvironmentVariables(location);

        location = location.Trim(' ', '\n', '\'', '"');

        if (Uri.TryCreate(location, UriKind.Absolute, out var uri) &&
            (location.Contains("://", StringComparison.Ordinal) || location.StartsWith("file:", StringComparison.OrdinalIgnoreCase)))
        {
            return uri.IsFile ? NormalizeFileSystemPath(uri.LocalPath) : location;
        }

        if (location.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
            return location;

        if (location.StartsWith("::", StringComparison.Ordinal))
            location = $"shell:{location}";

        else if (location.StartsWith("{", StringComparison.Ordinal))
            location = $"shell:::{location}";

        return NormalizeFileSystemPath(location);
    }

    private static string NormalizeFileSystemPath(string location)
    {
        location = location.Replace('/', '\\');
        var root = Path.GetPathRoot(location) ?? string.Empty;
        var trimmed = location.TrimEnd('\\');
        return trimmed.Length < root.Length && (root.Length == 1 || root.EndsWith(":\\", StringComparison.Ordinal))
            ? root
            : trimmed;
    }
    public static string GetExecutablePath()
    {
        var processName = Process.GetCurrentProcess().MainModule?.FileName;
        return processName is { Length: > 0 } ? processName : $"{AppDomain.CurrentDomain.FriendlyName}.exe";
    }
}
