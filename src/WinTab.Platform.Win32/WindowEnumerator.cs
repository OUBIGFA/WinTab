using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using WinTab.Core;

namespace WinTab.Platform.Win32;

public static class WindowEnumerator
{
    public static IReadOnlyList<WindowInfo> EnumerateTopLevelWindows(bool includeInvisible = false)
    {
        var results = new List<WindowInfo>();
        EnumWindows((hWnd, _) =>
        {
            var isVisible = IsWindowVisible(hWnd);
            if (!includeInvisible && !isVisible)
            {
                return true;
            }

            var title = GetWindowTitle(hWnd);
            var className = GetWindowClassName(hWnd);

            GetWindowThreadProcessId(hWnd, out var processId);
            var processName = string.Empty;
            if (processId != 0)
            {
                try
                {
                    processName = Process.GetProcessById((int)processId).ProcessName;
                }
                catch
                {
                    processName = string.Empty;
                }
            }

            results.Add(new WindowInfo(
                hWnd,
                title,
                processName,
                (int)processId,
                className,
                isVisible));

            return true;
        }, IntPtr.Zero);

        return results;
    }

    public static WindowInfo? TryGetWindowInfo(IntPtr hWnd, bool requireVisible = true)
    {
        if (hWnd == IntPtr.Zero)
        {
            return null;
        }

        var isVisible = IsWindowVisible(hWnd);
        if (requireVisible && !isVisible)
        {
            return null;
        }

        var title = GetWindowTitle(hWnd);
        var className = GetWindowClassName(hWnd);

        GetWindowThreadProcessId(hWnd, out var processId);
        var processName = string.Empty;
        if (processId != 0)
        {
            try
            {
                processName = Process.GetProcessById((int)processId).ProcessName;
            }
            catch
            {
                processName = string.Empty;
            }
        }

        return new WindowInfo(
            hWnd,
            title,
            processName,
            (int)processId,
            className,
            isVisible);
    }

    private static string GetWindowTitle(IntPtr hWnd)
    {
        var length = GetWindowTextLength(hWnd);
        if (length <= 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder(length + 1);
        _ = GetWindowText(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    private static string GetWindowClassName(IntPtr hWnd)
    {
        var builder = new StringBuilder(256);
        _ = GetClassName(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
