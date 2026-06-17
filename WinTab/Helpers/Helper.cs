using System;
using System.Linq;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using System.Diagnostics;
using System.ComponentModel;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using WinTab.Interop;
using WinTab.WinAPI;
using H.Hooks;

namespace WinTab.Helpers;

public static class Helper
{
    private static int _lastCtrlShiftCheckAt;
    private static bool _lastCtrlShiftCheckValue;
    private const int SM_XVIRTUALSCREEN = 76;
    private const int SM_YVIRTUALSCREEN = 77;
    private const int SM_CXVIRTUALSCREEN = 78;
    private const int SM_CYVIRTUALSCREEN = 79;
    private const int OffscreenRestoreMargin = 120;
    public static readonly ConcurrentDictionary<nint, RECT?> HiddenWindows = new();

    public static async Task DoDelayedBackgroundAsync(Action action, int delayMs = 2_000, CancellationToken cancellationToken = default)
    {
        await Task.Delay(delayMs, cancellationToken).ConfigureAwait(false);
        action();
    }
    public static async Task DoDelayedBackgroundAsync(Func<Task> action, int delayMs = 2_000, CancellationToken cancellationToken = default)
    {
        await Task.Delay(delayMs, cancellationToken).ConfigureAwait(false);
        await action().ConfigureAwait(false);
    }
    public static async Task<T> DoDelayedBackgroundAsync<T>(Func<Task<T>> action, int delayMs = 2_000, CancellationToken cancellationToken = default)
    {
        await Task.Delay(delayMs, cancellationToken).ConfigureAwait(false);
        return await action().ConfigureAwait(false);
    }

    public static T DoUntilNotDefault<T>(Func<T> action, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        return DoUntilCondition(
            action,
            result => !EqualityComparer<T?>.Default.Equals(result, default),
            timeMs,
            sleepMs,
            cancellationToken);
    }
    public static void DoUntilTimeEnd(Action action, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        DoUntilCondition(action, static () => false, timeMs, sleepMs, cancellationToken);
    }
    public static void DoUntilCondition(Action action, Func<bool> predicate, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        var startTicks = Stopwatch.GetTimestamp();

        while (!cancellationToken.IsCancellationRequested && !IsTimeUp(startTicks, timeMs))
        {
            action();
            if (predicate())
                return;

            Thread.Sleep(sleepMs);
        }
    }
    public static T DoUntilCondition<T>(Func<T> action, Predicate<T> predicate, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        var startTicks = Stopwatch.GetTimestamp();

        while (!cancellationToken.IsCancellationRequested && !IsTimeUp(startTicks, timeMs))
        {
            var result = action();
            if (predicate(result))
                return result;

            Thread.Sleep(sleepMs);
        }

        return action();
    }
    public static void DoIfCondition(Action action, Func<bool> predicate, bool justOnce = false, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        var startTicks = Stopwatch.GetTimestamp();

        while (!cancellationToken.IsCancellationRequested && !IsTimeUp(startTicks, timeMs))
        {
            if (predicate())
            {
                action();

                if (justOnce) return;
            }
            Thread.Sleep(sleepMs);
        }
    }
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
    public static Task DoUntilTimeEndAsync(Func<Task> action, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        return DoUntilConditionAsync(action, static () => false, timeMs, sleepMs, cancellationToken);
    }
    public static async Task DoUntilConditionAsync(Func<Task> action, Func<bool> predicate, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        var startTicks = Stopwatch.GetTimestamp();

        while (!cancellationToken.IsCancellationRequested && !IsTimeUp(startTicks, timeMs))
        {
            await action().ConfigureAwait(false);
            if (predicate())
                return;

            await Task.Delay(sleepMs, cancellationToken).ConfigureAwait(false);
        }
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
    public static async Task DoIfConditionAsync(Func<Task> action, Func<bool> predicate, bool justOnce = false, int timeMs = 500, int sleepMs = 20, CancellationToken cancellationToken = default)
    {
        var startTicks = Stopwatch.GetTimestamp();

        while (!cancellationToken.IsCancellationRequested && !IsTimeUp(startTicks, timeMs))
        {
            if (predicate())
            {
                await action().ConfigureAwait(false);

                if (justOnce) return;
            }
            await Task.Delay(sleepMs, cancellationToken).ConfigureAwait(false);
        }
    }

    public static bool IsTimeUp(long startTicks, int timeMs)
    {

#if NET7_0_OR_GREATER
        var elapsedTime = Stopwatch.GetElapsedTime(startTicks);
#else
        var elapsedTime = GetElapsedTime(startTicks);
#endif

        return elapsedTime.TotalMilliseconds >= timeMs;
    }
    public static TimeSpan GetElapsedTime(long startTicks)
    {
        var tickFrequency = (double)10_000_000 / Stopwatch.Frequency;
        return new TimeSpan((long)((Stopwatch.GetTimestamp() - startTicks) * tickFrequency));
    }
    public static T Clamp<T>(this T val, T min, T max) where T : IComparable<T>
    {
        if (val.CompareTo(min) < 0) return min;
        if (val.CompareTo(max) > 0) return max;
        return val;
    }

    public static Icon? GetIcon() => Icon.ExtractAssociatedIcon(GetExecutablePath());
    public static string GetEnumDescription(Enum value)
    {
        var fieldInfo = value.GetType().GetField(value.ToString());
        return fieldInfo?.GetCustomAttribute<DescriptionAttribute>()?.Description ?? value.ToString();
    }
    public static string HotKeysToString(this IEnumerable<Key> keys, bool isDoubleClick = false)
    {
        var text = string.Join(" + ", keys.Select(k => k.ToDisplayString()));
        if (isDoubleClick) text += "_DBL";
        return text;
    }
    public static string ToDisplayString(this Key key)
    {
        return key switch
        {
            Key.Add => "+",
            Key.Subtract => "-",
            Key.Multiply => "*",
            Key.Divide => "/",
            Key.OemPlus => "+",
            Key.OemMinus => "-",
            Key.OemComma => ",",
            Key.Decimal or Key.OemPeriod => "DOT",
            Key.Oem1 => ";",
            Key.Oem2 => "/",
            Key.Oem3 => "Tilde",
            Key.Oem4 => "[",
            Key.Oem5 => "\\",
            Key.Oem6 => "]",
            Key.Oem7 => "Quote",
            Key.Escape => "ESC",
            Key.CapsLock => "CAPS",
            Key.PageUp => "PgUp",
            Key.PageDown => "PgDn",
            Key.PrintScreen => "PrtSc",

            >= Key.NumPad0 and <= Key.NumPad9 => key.ToString().Replace("NumPad", "Num"),
            >= Key.D0 and <= Key.D9 => key.ToString().Replace("D", ""),

            // Mouse buttons
            Key.MouseLeft or Key.LButton => "LMB",
            Key.MouseRight or Key.RButton => "RMB",
            Key.MouseMiddle or Key.MButton => "MMB",
            Key.MouseXButton1 => "X1",
            Key.MouseXButton2 => "X2",

            // Default case
            _ => key.ToFixedString().Replace("Button", "")
                .Replace("Mouse", "")
                .Replace("Key", "")
                .Trim()
        };
    }

    public static bool IsExplorerEmptySpace(Point point)
    {
        var hr = WinApi.AccessibleObjectFromPoint(point, out var accObj, out var childId);
        if (hr != 0 || childId is not 0 || accObj is not IAccessible accessible) return false;

        var role = accessible.get_accRole(0);
        return role is 0x21; //IAccessible.Role:list (ROLE_SYSTEM_LIST 0x21)
    }
    public static bool IsFileExplorerTab(nint tab)
    {
        return tab != 0 && WinApi.IsWindowHasClassName(tab, "ShellTabWindowClass");
    }
    public static bool IsFileExplorerWindow(nint window)
    {
        return window != 0 && WinApi.IsWindowHasClassName(window, "CabinetWClass");
    }
    public static bool IsFileExplorerForeground(out nint foregroundWindow)
    {
        foregroundWindow = WinApi.GetForegroundWindow();
        return IsFileExplorerWindow(foregroundWindow);
    }
    public static nint GetAnotherExplorerWindow(nint currentWindow)
    {
        return currentWindow == 0
            ? WinApi.FindWindow("CabinetWClass", null)
            : GetAllExplorerWindows()
                .FirstOrDefault(window => window != currentWindow);
    }
    public static Task<nint> ListenForNewExplorerWindowAsync(IReadOnlyCollection<nint> currentWindows, int searchTimeMs = 1000)
    {
        var knownWindows = CreateKnownHandleSet(currentWindows);
        return DoUntilNotDefaultAsync(() =>
                GetAllExplorerWindows()
                    .FirstOrDefault(window => IsUnknownHandle(window, knownWindows)),
            searchTimeMs);
    }

    public static nint ListenForNewExplorerTab(IReadOnlyCollection<nint> currentTabs, int searchTimeMs = 1000)
    {
        var knownTabs = CreateKnownHandleSet(currentTabs);
        return DoUntilNotDefault(() =>
                GetAllExplorerTabs()
                    .FirstOrDefault(tab => IsUnknownHandle(tab, knownTabs)),
            searchTimeMs);
    }
    public static Task<nint> ListenForNewExplorerTabAsync(IReadOnlyCollection<nint> currentTabs, int searchTimeMs = 1000)
    {
        var knownTabs = CreateKnownHandleSet(currentTabs);
        return DoUntilNotDefaultAsync(() =>
                GetAllExplorerTabs()
                    .FirstOrDefault(tab => IsUnknownHandle(tab, knownTabs)),
            searchTimeMs);
    }
    public static Task<nint> ListenForNewExplorerTabAsync(nint window, IReadOnlyCollection<nint> currentTabs, int searchTimeMs = 1000)
    {
        var knownTabs = CreateKnownHandleSet(currentTabs);
        return DoUntilNotDefaultAsync(() =>
                GetAllExplorerTabs(window)
                    .FirstOrDefault(tab => IsUnknownHandle(tab, knownTabs)),
            searchTimeMs);
    }
    private static HashSet<nint>? CreateKnownHandleSet(IReadOnlyCollection<nint> handles)
    {
        return handles.Count == 0 ? null : new HashSet<nint>(handles);
    }
    private static bool IsUnknownHandle(nint handle, HashSet<nint>? knownHandles)
    {
        return knownHandles == null || !knownHandles.Contains(handle);
    }
    public static List<nint> GetAllExplorerTabs()
    {
        var tabs = new List<nint>();

        foreach (var window in GetAllExplorerWindows())
            tabs.AddRange(GetAllExplorerTabs(window));

        return tabs;
    }
    public static IEnumerable<nint> GetAllExplorerTabs(nint window)
    {
        return WinApi.FindAllWindowsEx("ShellTabWindowClass", window);
    }
    public static IEnumerable<nint> GetAllExplorerWindows()
    {
        return WinApi.FindAllWindowsEx("CabinetWClass");
    }
    public static Process? GetMainExplorerProcess()
    {
        Process? best = null;
        var windowsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        var expectedPath = System.IO.Path.Combine(windowsFolder, "explorer.exe");
        var bestStart = DateTime.MaxValue;

        foreach (var hWnd in WinApi.FindAllWindowsEx("Shell_TrayWnd")) // Taskbar
        {
            if (WinApi.GetWindowThreadProcessId(hWnd, out var pid) <= 0) continue;
        
            var processPath = WinApi.GetProcessPath((int)pid);
            if (!string.Equals(processPath, expectedPath, StringComparison.OrdinalIgnoreCase))
                continue;

            try
            {
                // Pick the earliest start
                var proc = Process.GetProcessById((int)pid);
                if (proc.StartTime < bestStart)
                {
                    bestStart = proc.StartTime;
                    best = proc;
                }
            }
            catch { /* The Process might have terminated */ }
        }
        return best;
    }
    
    public static void UpdateWindowLayered(nint hWnd, bool remove)
    {
        var exStyle = WinApi.GetWindowLong(hWnd, WinApi.GWL_EXSTYLE);
        var isLayered = (exStyle & WinApi.WS_EX_LAYERED) != 0;
        
        if (remove && isLayered) // Remove
            WinApi.SetWindowLong(hWnd, WinApi.GWL_EXSTYLE, exStyle & ~WinApi.WS_EX_LAYERED);
        
        if (!remove && !isLayered) // Add
            WinApi.SetWindowLong(hWnd, WinApi.GWL_EXSTYLE, exStyle | WinApi.WS_EX_LAYERED);
    }
    public static void HideWindow(nint hWnd, bool keepTheme = false)
    {
        var originalPos = HiddenWindows.GetOrAdd(hWnd, static (hWnd, keepTheme) =>
        {
            if (!keepTheme)
                return null;

            return WinApi.GetWindowRect(hWnd, out var originalPos) ? originalPos : null;
        }, keepTheme);

        if (!keepTheme)
        {
            UpdateWindowLayered(hWnd, remove: false);
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 0, WinApi.LWA_ALPHA);
            return;
        }

        if (originalPos == null && WinApi.GetWindowRect(hWnd, out var currentPos))
        {
            originalPos = currentPos;
            HiddenWindows[hWnd] = currentPos;
        }

        const uint flags = WinApi.SWP_HIDEWINDOW | WinApi.SWP_NOSIZE | WinApi.SWP_NOZORDER | WinApi.SWP_NOACTIVATE | WinApi.SWP_FRAMECHANGED;
        WinApi.SetWindowPos(hWnd, 0, -32_000, -32_000, 0, 0, flags);
    }
    public static bool ShowWindow(nint hWnd, bool removeCache)
    {
        var restored = RestoreHiddenExplorerWindow(hWnd, removeCache, removeLayeredStyle: false);
        if (restored)
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 255, WinApi.LWA_ALPHA);

        return restored;
    }
    public static int RestoreHiddenExplorerWindows(bool removeLayeredStyle = true)
    {
        var restored = 0;
        var candidates = GetAllExplorerWindows()
            .Concat(HiddenWindows.Keys)
            .Distinct()
            .ToArray();

        foreach (var hWnd in candidates)
        {
            if (RestoreHiddenExplorerWindow(hWnd, removeCache: true, removeLayeredStyle))
                restored++;
        }

        return restored;
    }
    public static bool RestoreHiddenExplorerWindow(nint hWnd, bool removeCache = true, bool removeLayeredStyle = true)
    {
        if (hWnd == 0 || !IsFileExplorerWindow(hWnd))
        {
            if (removeCache)
                HiddenWindows.TryRemove(hWnd, out _);

            return false;
        }

        var hasCache = HiddenWindows.TryGetValue(hWnd, out var originalPos);
        var exStyle = WinApi.GetWindowLong(hWnd, WinApi.GWL_EXSTYLE);
        var isLayered = (exStyle & WinApi.WS_EX_LAYERED) != 0;
        var alpha = (byte)255;
        var alphaFlags = 0u;
        var isTransparent = isLayered &&
                            WinApi.GetLayeredWindowAttributes(hWnd, out _, out alpha, out alphaFlags) &&
                            (alphaFlags & (uint)WinApi.LWA_ALPHA) != 0 &&
                            alpha == 0;
        var isVisible = WinApi.IsWindowVisible(hWnd);
        var hasRect = WinApi.GetWindowRect(hWnd, out var rect);
        var isOffscreen = hasRect && IsExplorerWindowOffscreen(rect);

        if (!hasCache && !isTransparent && isVisible && !isOffscreen)
            return false;

        if (removeCache)
            HiddenWindows.TryRemove(hWnd, out _);

        const uint showFlags = WinApi.SWP_SHOWWINDOW | WinApi.SWP_NOSIZE | WinApi.SWP_NOZORDER | WinApi.SWP_NOACTIVATE | WinApi.SWP_FRAMECHANGED;
        if (originalPos != null)
        {
            WinApi.SetWindowPos(hWnd, 0, originalPos.Value.Left, originalPos.Value.Top, 0, 0, showFlags);
        }
        else if (isOffscreen)
        {
            var x = WinApi.GetSystemMetrics(SM_XVIRTUALSCREEN) + OffscreenRestoreMargin;
            var y = WinApi.GetSystemMetrics(SM_YVIRTUALSCREEN) + OffscreenRestoreMargin;
            WinApi.SetWindowPos(hWnd, 0, x, y, 0, 0, showFlags);
        }
        else
        {
            WinApi.ShowWindow(hWnd, WinApi.SW_SHOWNOACTIVATE);
            if (hasRect)
                WinApi.SetWindowPos(hWnd, 0, rect.Left, rect.Top, 0, 0, showFlags);
        }

        if (isLayered)
        {
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 255, WinApi.LWA_ALPHA);
            if (removeLayeredStyle)
                UpdateWindowLayered(hWnd, remove: true);
        }

        return true;
    }
    private static bool IsExplorerWindowOffscreen(RECT rect)
    {
        var width = rect.Right - rect.Left;
        var height = rect.Bottom - rect.Top;
        if (width < 80 || height < 80)
            return false;

        var virtualLeft = WinApi.GetSystemMetrics(SM_XVIRTUALSCREEN);
        var virtualTop = WinApi.GetSystemMetrics(SM_YVIRTUALSCREEN);
        var virtualRight = virtualLeft + WinApi.GetSystemMetrics(SM_CXVIRTUALSCREEN);
        var virtualBottom = virtualTop + WinApi.GetSystemMetrics(SM_CYVIRTUALSCREEN);

        return rect.Right < virtualLeft - OffscreenRestoreMargin ||
               rect.Bottom < virtualTop - OffscreenRestoreMargin ||
               rect.Left > virtualRight + OffscreenRestoreMargin ||
               rect.Top > virtualBottom + OffscreenRestoreMargin;
    }

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
