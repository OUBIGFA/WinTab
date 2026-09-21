using System;
using System.Runtime.InteropServices;
using WinTab.Interop;

namespace WinTab.Helpers;

/// <summary>
/// Removes and restores a window's taskbar button. A merge source is on screen while it is concealed, so the
/// taskbar would otherwise count it as a second window until Explorer has closed it.
/// </summary>
internal static class TaskbarButton
{
    /// <summary>Removes the window's button. The button stays removed even if the window is shown or activated later.</summary>
    public static bool Remove(nint hWnd) => Invoke(hWnd, static (list, handle) => list.DeleteTab(handle), "remove");

    /// <summary>Puts the window's button back; the taskbar shows it once the window is visible.</summary>
    public static bool Add(nint hWnd) => Invoke(hWnd, static (list, handle) => list.AddTab(handle), "restore");

    private static bool Invoke(nint hWnd, Func<ITaskbarList, nint, int> action, string operation)
    {
        if (hWnd == 0)
            return false;
        ITaskbarList? list = null;
        try
        {
            list = (ITaskbarList)new TaskbarList();
            var result = list.HrInit();
            if (result == 0)
                result = action(list, hWnd);
            if (result == 0)
                return true;
            System.Diagnostics.Trace.TraceWarning($"Taskbar button {operation} failed for {hWnd}: 0x{result:X8}");
            return false;
        }
        catch (Exception exception) when (exception is COMException or InvalidCastException or NotSupportedException)
        {
            System.Diagnostics.Trace.TraceWarning($"Taskbar button {operation} unavailable for {hWnd}: {exception.GetType().Name}");
            return false;
        }
        finally
        {
            if (list != null)
                Marshal.ReleaseComObject(list);
        }
    }
}
