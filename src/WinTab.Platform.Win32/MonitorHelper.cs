using System.Runtime.InteropServices;

namespace WinTab.Platform.Win32;

/// <summary>
/// Multi-monitor support utilities.
/// Retrieves monitor work areas and bounds using Win32 APIs.
/// </summary>
public static class MonitorHelper
{
    /// <summary>
    /// Gets the work area (usable area excluding taskbar) of the monitor
    /// that contains the largest portion of the specified window.
    /// </summary>
    /// <param name="hWnd">Handle of the window to locate.</param>
    /// <returns>The work area rectangle of the monitor.</returns>
    public static NativeStructs.RECT GetMonitorBounds(IntPtr hWnd)
    {
        IntPtr hMonitor = NativeMethods.MonitorFromWindow(
            hWnd, NativeConstants.MONITOR_DEFAULTTONEAREST);

        if (hMonitor == IntPtr.Zero)
            return default;

        var monitorInfo = new NativeStructs.MONITORINFOEX();
        monitorInfo.cbSize = (uint)Marshal.SizeOf<NativeStructs.MONITORINFOEX>();

        if (NativeMethods.GetMonitorInfo(hMonitor, ref monitorInfo))
            return monitorInfo.rcWork;

        return default;
    }

    /// <summary>
    /// Gets the full monitor rectangle (including taskbar area) of the monitor
    /// that contains the largest portion of the specified window.
    /// </summary>
    /// <param name="hWnd">Handle of the window to locate.</param>
    /// <returns>The full monitor rectangle.</returns>
    public static NativeStructs.RECT GetMonitorFullBounds(IntPtr hWnd)
    {
        IntPtr hMonitor = NativeMethods.MonitorFromWindow(
            hWnd, NativeConstants.MONITOR_DEFAULTTONEAREST);

        if (hMonitor == IntPtr.Zero)
            return default;

        var monitorInfo = new NativeStructs.MONITORINFOEX();
        monitorInfo.cbSize = (uint)Marshal.SizeOf<NativeStructs.MONITORINFOEX>();

        if (NativeMethods.GetMonitorInfo(hMonitor, ref monitorInfo))
            return monitorInfo.rcMonitor;

        return default;
    }

    /// <summary>
    /// Enumerates all connected monitors and returns their work area rectangles.
    /// </summary>
    /// <returns>A list of work area rectangles for each monitor.</returns>
    public static List<NativeStructs.RECT> GetAllMonitors()
    {
        var monitors = new List<NativeStructs.RECT>();

        NativeMethods.EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero,
            (IntPtr hMonitor, IntPtr hdcMonitor, ref NativeStructs.RECT lprcMonitor, IntPtr dwData) =>
            {
                var monitorInfo = new NativeStructs.MONITORINFOEX();
                monitorInfo.cbSize = (uint)Marshal.SizeOf<NativeStructs.MONITORINFOEX>();

                if (NativeMethods.GetMonitorInfo(hMonitor, ref monitorInfo))
                    monitors.Add(monitorInfo.rcWork);

                return true; // Continue enumeration.
            },
            IntPtr.Zero);

        return monitors;
    }

    /// <summary>
    /// Determines whether a point is within any monitor's bounds.
    /// </summary>
    /// <param name="x">Screen X coordinate.</param>
    /// <param name="y">Screen Y coordinate.</param>
    /// <returns>True if the point is within a monitor's work area.</returns>
    public static bool IsPointOnScreen(int x, int y)
    {
        foreach (var rect in GetAllMonitors())
        {
            if (x >= rect.Left && x <= rect.Right && y >= rect.Top && y <= rect.Bottom)
                return true;
        }
        return false;
    }
}
