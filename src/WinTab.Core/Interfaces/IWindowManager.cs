using WinTab.Core.Models;

namespace WinTab.Core.Interfaces;

public interface IWindowManager
{
    bool Show(IntPtr hWnd);
    bool Hide(IntPtr hWnd);
    bool Minimize(IntPtr hWnd);
    bool Restore(IntPtr hWnd);
    bool Close(IntPtr hWnd);
    bool BringToFront(IntPtr hWnd);
    bool IsAlive(IntPtr hWnd);
    bool IsVisible(IntPtr hWnd);
    bool SetBounds(IntPtr hWnd, int x, int y, int width, int height);
    (int X, int Y, int Width, int Height) GetBounds(IntPtr hWnd);
    IReadOnlyList<WindowInfo> EnumerateTopLevelWindows(bool includeInvisible = false);
    WindowInfo? GetWindowInfo(IntPtr hWnd);
}
