using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections.Generic;
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
    public static async Task<T> DoUntilConditionAsync<T>(Func<T> action, Predicate<T> predicate, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        var startTicks = Stopwatch.GetTimestamp();

        while (!cancellationToken.IsCancellationRequested && !IsTimeUp(startTicks, timeMs))
        {
            var result = action();
            if (predicate(result))
                return result;

            await Task.Delay(sleepMs, cancellationToken).ConfigureAwait(false);
        }

        return action();
    }
    public static async Task<T> DoUntilConditionAsync<T>(Func<Task<T>> action, Predicate<T> predicate, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        var startTicks = Stopwatch.GetTimestamp();

        while (!cancellationToken.IsCancellationRequested && !IsTimeUp(startTicks, timeMs))
        {
            var result = await action().ConfigureAwait(false);
            if (predicate(result))
                return result;

            await Task.Delay(sleepMs, cancellationToken).ConfigureAwait(false);
        }

        return await action().ConfigureAwait(false);
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
        bool isUnc = location.StartsWith("\\\\") || location.StartsWith("//");
        if (isUnc)
        {
            location = location.TrimEnd('/', '\\');
        }
        else
        {
            location = location.Trim('/', '\\');
        }

        if (Uri.TryCreate(location, UriKind.Absolute, out var uri) &&
            (location.Contains("://", StringComparison.Ordinal) || location.StartsWith("file:", StringComparison.OrdinalIgnoreCase)))
        {
            return uri.IsFile ? uri.LocalPath.TrimEnd('\\', '/') : location;
        }

        if (location.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
            return location;

        if (location.StartsWith("::", StringComparison.Ordinal))
            location = $"shell:{location}";

        else if (location.StartsWith("{", StringComparison.Ordinal))
            location = $"shell:::{location}";

        return location.Replace('/', '\\');
    }
    public static string GetExecutablePath()
    {
        var processName = Process.GetCurrentProcess().MainModule?.FileName;
        return processName is { Length: > 0 } ? processName : $"{AppDomain.CurrentDomain.FriendlyName}.exe";
    }
}
