using WinTab.Core.Interfaces;
using WinTab.Core.Models;

namespace WinTab.Platform.Win32;

/// <summary>
/// Win32 implementation of <see cref="IWindowManager"/>.
/// Provides window manipulation operations through native P/Invoke calls.
/// </summary>
public sealed class WindowManager : IWindowManager
{
    /// <inheritdoc />
    public bool Show(IntPtr hWnd)
    {
        if (!NativeMethods.IsWindow(hWnd))
            return false;

        return NativeMethods.ShowWindow(hWnd, NativeConstants.SW_SHOW);
    }

    /// <inheritdoc />
    public bool Hide(IntPtr hWnd)
    {
        if (!NativeMethods.IsWindow(hWnd))
            return false;

        return NativeMethods.ShowWindow(hWnd, NativeConstants.SW_HIDE);
    }

    /// <inheritdoc />
    public bool Minimize(IntPtr hWnd)
    {
        if (!NativeMethods.IsWindow(hWnd))
            return false;

        return NativeMethods.ShowWindow(hWnd, NativeConstants.SW_MINIMIZE);
    }

    /// <inheritdoc />
    public bool Restore(IntPtr hWnd)
    {
        if (!NativeMethods.IsWindow(hWnd))
            return false;

        return NativeMethods.ShowWindow(hWnd, NativeConstants.SW_RESTORE);
    }

    /// <inheritdoc />
    public bool Close(IntPtr hWnd)
    {
        if (!NativeMethods.IsWindow(hWnd))
            return false;

        return NativeMethods.PostMessage(hWnd, (uint)NativeConstants.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }

    /// <inheritdoc />
    public bool BringToFront(IntPtr hWnd)
    {
        if (!NativeMethods.IsWindow(hWnd))
            return false;

        // If the window is minimized, restore it first.
        var placement = new NativeStructs.WINDOWPLACEMENT();
        placement.length = (uint)System.Runtime.InteropServices.Marshal.SizeOf<NativeStructs.WINDOWPLACEMENT>();
        if (NativeMethods.GetWindowPlacement(hWnd, ref placement))
        {
            if (placement.showCmd == (uint)NativeConstants.SW_SHOWMINIMIZED)
                NativeMethods.ShowWindow(hWnd, NativeConstants.SW_RESTORE);
        }

        // Show the window, then bring it to the foreground.
        NativeMethods.ShowWindow(hWnd, NativeConstants.SW_SHOW);
        return NativeMethods.SetForegroundWindow(hWnd);
    }

    /// <inheritdoc />
    public bool IsAlive(IntPtr hWnd)
    {
        return NativeMethods.IsWindow(hWnd);
    }

    /// <inheritdoc />
    public bool IsVisible(IntPtr hWnd)
    {
        return NativeMethods.IsWindowVisible(hWnd);
    }

    /// <inheritdoc />
    public bool SetBounds(IntPtr hWnd, int x, int y, int width, int height)
    {
        if (!NativeMethods.IsWindow(hWnd))
            return false;

        return NativeMethods.MoveWindow(hWnd, x, y, width, height, true);
    }

    /// <inheritdoc />
    public (int X, int Y, int Width, int Height) GetBounds(IntPtr hWnd)
    {
        if (!NativeMethods.IsWindow(hWnd) ||
            !NativeMethods.GetWindowRect(hWnd, out NativeStructs.RECT rect))
        {
            return (0, 0, 0, 0);
        }

        return (rect.Left, rect.Top, rect.Width, rect.Height);
    }

    /// <inheritdoc />
    public IReadOnlyList<WindowInfo> EnumerateTopLevelWindows(bool includeInvisible = false)
    {
        return WindowEnumerator.EnumerateTopLevelWindows(includeInvisible);
    }

    /// <inheritdoc />
    public WindowInfo? GetWindowInfo(IntPtr hWnd)
    {
        return WindowEnumerator.GetWindowInfo(hWnd);
    }
}
