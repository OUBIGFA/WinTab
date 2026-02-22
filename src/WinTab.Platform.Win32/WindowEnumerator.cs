using System.Diagnostics;
using System.Text;
using WinTab.Core.Models;

namespace WinTab.Platform.Win32;

/// <summary>
/// Enumerates top-level windows and gathers <see cref="WindowInfo"/> records.
/// </summary>
public static class WindowEnumerator
{
    /// <summary>
    /// Enumerates all top-level windows that qualify as "application windows".
    /// </summary>
    /// <param name="includeInvisible">If true, invisible windows are also returned.</param>
    /// <returns>A read-only list of <see cref="WindowInfo"/> for each qualifying window.</returns>
    public static IReadOnlyList<WindowInfo> EnumerateTopLevelWindows(bool includeInvisible = false)
    {
        var windows = new List<WindowInfo>();

        NativeMethods.EnumWindows((hWnd, _) =>
        {
            if (!includeInvisible && !NativeMethods.IsWindowVisible(hWnd))
                return true;

            // Skip windows with no title.
            int titleLength = NativeMethods.GetWindowTextLength(hWnd);
            if (titleLength == 0)
                return true;

            // Read styles — filter out child windows and tool windows.
            long style = NativeMethods.GetWindowLongPtr(hWnd, NativeConstants.GWL_STYLE).ToInt64();
            long exStyle = NativeMethods.GetWindowLongPtr(hWnd, NativeConstants.GWL_EXSTYLE).ToInt64();

            if ((style & NativeConstants.WS_CHILD) != 0)
                return true;

            if ((exStyle & NativeConstants.WS_EX_TOOLWINDOW) != 0)
                return true;

            var info = BuildWindowInfo(hWnd);
            if (info is not null)
                windows.Add(info);

            return true;
        }, IntPtr.Zero);

        return windows;
    }

    /// <summary>
    /// Retrieves <see cref="WindowInfo"/> for a specific window handle.
    /// </summary>
    /// <param name="hWnd">The window handle.</param>
    /// <returns>A <see cref="WindowInfo"/> record, or null if the handle is invalid.</returns>
    public static WindowInfo? GetWindowInfo(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero || !NativeMethods.IsWindow(hWnd))
            return null;

        return BuildWindowInfo(hWnd);
    }

    // ─── Private Helpers ────────────────────────────────────────────────────

    private static WindowInfo? BuildWindowInfo(IntPtr hWnd)
    {
        // Title
        int titleLen = NativeMethods.GetWindowTextLength(hWnd);
        var titleBuilder = new StringBuilder(titleLen + 1);
        NativeMethods.GetWindowText(hWnd, titleBuilder, titleBuilder.Capacity);
        string title = titleBuilder.ToString();

        // Class name
        var classBuilder = new StringBuilder(256);
        NativeMethods.GetClassName(hWnd, classBuilder, classBuilder.Capacity);
        string className = classBuilder.ToString();

        // Process info
        NativeMethods.GetWindowThreadProcessId(hWnd, out uint pid);
        string processName = string.Empty;
        string? processPath = null;
        int processId = (int)pid;

        if (pid != 0)
        {
            try
            {
                using var process = Process.GetProcessById(processId);
                processName = process.ProcessName;
            }
            catch (ArgumentException)
            {
                // Process has already exited.
            }
            catch (InvalidOperationException)
            {
                // Process has already exited.
            }

            // Attempt to get process path via QueryFullProcessImageName for elevated processes.
            processPath = GetProcessPath(pid);
        }

        bool isVisible = NativeMethods.IsWindowVisible(hWnd);

        return new WindowInfo(
            Handle: hWnd,
            Title: title,
            ProcessName: processName,
            ProcessId: processId,
            ClassName: className,
            IsVisible: isVisible,
            ProcessPath: processPath);
    }

    /// <summary>
    /// Retrieves the full executable path for a process using QueryFullProcessImageName.
    /// Falls back gracefully when access is denied.
    /// </summary>
    private static string? GetProcessPath(uint pid)
    {
        IntPtr hProcess = NativeMethods.OpenProcess(
            NativeConstants.PROCESS_QUERY_LIMITED_INFORMATION, false, pid);

        if (hProcess == IntPtr.Zero)
            return null;

        try
        {
            var buffer = new StringBuilder(1024);
            uint size = (uint)buffer.Capacity;

            if (NativeMethods.QueryFullProcessImageName(hProcess, 0, buffer, ref size))
                return buffer.ToString();

            return null;
        }
        finally
        {
            NativeMethods.CloseHandle(hProcess);
        }
    }
}
