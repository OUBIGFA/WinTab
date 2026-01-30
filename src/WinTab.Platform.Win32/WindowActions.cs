using System.Runtime.InteropServices;

namespace WinTab.Platform.Win32;

public static class WindowActions
{
    public static bool BringToFront(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        _ = ShowWindow(hWnd, ShowWindowCommand.Restore);
        return SetForegroundWindow(hWnd);
    }

    public static bool Minimize(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        return ShowWindow(hWnd, ShowWindowCommand.Minimize);
    }

    public static bool Restore(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        return ShowWindow(hWnd, ShowWindowCommand.Restore);
    }

    public static bool Show(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        return ShowWindow(hWnd, ShowWindowCommand.ShowNormal);
    }

    public static bool Hide(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        return ShowWindow(hWnd, ShowWindowCommand.Hide);
    }

    public static bool IsAlive(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        return IsWindow(hWnd);
    }

    public static bool Close(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        return PostMessage(hWnd, WindowMessage.Close, IntPtr.Zero, IntPtr.Zero);
    }

    private enum ShowWindowCommand
    {
        Hide = 0,
        ShowNormal = 1,
        ShowMinimized = 2,
        ShowMaximized = 3,
        Restore = 9,
        Minimize = 6
    }

    private enum WindowMessage
    {
        Close = 0x0010
    }

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, ShowWindowCommand nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, WindowMessage msg, IntPtr wParam, IntPtr lParam);
}
