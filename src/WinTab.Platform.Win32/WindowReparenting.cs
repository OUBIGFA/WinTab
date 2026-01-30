using System.Runtime.InteropServices;

namespace WinTab.Platform.Win32;

public sealed class ReparentedWindow
{
    internal ReparentedWindow(IntPtr handle, IntPtr originalParent, int originalStyle, int originalExStyle, WindowRect originalRect)
    {
        Handle = handle;
        OriginalParent = originalParent;
        OriginalStyle = originalStyle;
        OriginalExStyle = originalExStyle;
        OriginalRect = originalRect;
    }

    public IntPtr Handle { get; }
    public IntPtr OriginalParent { get; }
    public int OriginalStyle { get; }
    public int OriginalExStyle { get; }
    internal WindowRect OriginalRect { get; }
}

[StructLayout(LayoutKind.Sequential)]
internal struct WindowRect
{
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;

    public int Width => Right - Left;
    public int Height => Bottom - Top;
}

public static class WindowReparenting
{
    public static ReparentedWindow AttachToHost(IntPtr child, IntPtr host)
    {
        var originalParent = GetParent(child);
        var originalStyle = GetWindowLong(child, GwlStyle);
        var originalExStyle = GetWindowLong(child, GwlExStyle);
        _ = GetWindowRect(child, out var originalRect);

        var newStyle = originalStyle;
        newStyle &= ~WsCaption;
        newStyle &= ~WsThickFrame;
        newStyle &= ~WsMinimizeBox;
        newStyle &= ~WsMaximizeBox;
        newStyle &= ~WsSysMenu;
        newStyle |= WsChild;

        SetWindowLong(child, GwlStyle, newStyle);
        SetParent(child, host);
        SetWindowPos(child, IntPtr.Zero, 0, 0, 0, 0, SwpNoZOrder | SwpNoSize | SwpShowWindow);

        return new ReparentedWindow(child, originalParent, originalStyle, originalExStyle, originalRect);
    }

    public static void ResizeToHost(IntPtr child, int width, int height)
    {
        if (child == IntPtr.Zero || width <= 0 || height <= 0)
        {
            return;
        }

        SetWindowPos(child, IntPtr.Zero, 0, 0, width, height, SwpNoZOrder);
    }

    public static void Detach(ReparentedWindow window)
    {
        SetParent(window.Handle, window.OriginalParent);
        SetWindowLong(window.Handle, GwlStyle, window.OriginalStyle);
        SetWindowLong(window.Handle, GwlExStyle, window.OriginalExStyle);

        var rect = window.OriginalRect;
        var width = rect.Width;
        var height = rect.Height;

        if (width > 0 && height > 0)
        {
            SetWindowPos(window.Handle, IntPtr.Zero, rect.Left, rect.Top, width, height,
                SwpNoZOrder | SwpFrameChanged | SwpShowWindow);
        }
        else
        {
            SetWindowPos(window.Handle, IntPtr.Zero, 0, 0, 0, 0,
                SwpNoMove | SwpNoSize | SwpNoZOrder | SwpFrameChanged | SwpShowWindow);
        }
    }

    private const int GwlStyle = -16;
    private const int GwlExStyle = -20;

    private const int WsChild = 0x40000000;
    private const int WsCaption = 0x00C00000;
    private const int WsThickFrame = 0x00040000;
    private const int WsMinimizeBox = 0x00020000;
    private const int WsMaximizeBox = 0x00010000;
    private const int WsSysMenu = 0x00080000;

    private const int SwpNoSize = 0x0001;
    private const int SwpNoMove = 0x0002;
    private const int SwpNoZOrder = 0x0004;
    private const int SwpShowWindow = 0x0040;
    private const int SwpFrameChanged = 0x0020;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetWindowRect(IntPtr hWnd, out WindowRect lpRect);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int x,
        int y,
        int cx,
        int cy,
        int uFlags);
}
