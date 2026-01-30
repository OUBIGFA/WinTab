using System.Runtime.InteropServices;

namespace WinTab.Platform.Win32;

public sealed class WindowEventWatcher : IDisposable
{
    public event EventHandler<IntPtr>? WindowShown;
    public event EventHandler<IntPtr>? WindowDestroyed;

    private readonly WinEventDelegate _callback;
    private IntPtr _showHook;
    private IntPtr _destroyHook;
    private bool _disposed;

    public WindowEventWatcher()
    {
        _callback = HandleWinEvent;

        _showHook = SetWinEventHook(
            EventObjectShow,
            EventObjectShow,
            IntPtr.Zero,
            _callback,
            0,
            0,
            WineventOutOfContext | WineventSkipOwnProcess);

        _destroyHook = SetWinEventHook(
            EventObjectDestroy,
            EventObjectDestroy,
            IntPtr.Zero,
            _callback,
            0,
            0,
            WineventOutOfContext | WineventSkipOwnProcess);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        if (_showHook != IntPtr.Zero)
        {
            _ = UnhookWinEvent(_showHook);
            _showHook = IntPtr.Zero;
        }

        if (_destroyHook != IntPtr.Zero)
        {
            _ = UnhookWinEvent(_destroyHook);
            _destroyHook = IntPtr.Zero;
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }

    private void HandleWinEvent(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr hwnd,
        int idObject,
        int idChild,
        uint eventThread,
        uint eventTime)
    {
        if (hwnd == IntPtr.Zero || idObject != ObjidWindow || idChild != 0)
        {
            return;
        }

        if (eventType == EventObjectShow)
        {
            if (!IsTopLevelWindow(hwnd) || !IsWindowVisible(hwnd))
            {
                return;
            }

            WindowShown?.Invoke(this, hwnd);
            return;
        }

        if (eventType == EventObjectDestroy)
        {
            if (!IsTopLevelWindow(hwnd))
            {
                return;
            }

            WindowDestroyed?.Invoke(this, hwnd);
        }
    }

    private static bool IsTopLevelWindow(IntPtr hwnd)
    {
        var root = GetAncestor(hwnd, GaRoot);
        return root != IntPtr.Zero && root == hwnd;
    }

    private const uint EventObjectDestroy = 0x8001;
    private const uint EventObjectShow = 0x8002;
    private const int ObjidWindow = 0;

    private const uint WineventOutOfContext = 0x0000;
    private const uint WineventSkipOwnProcess = 0x0002;

    private const uint GaRoot = 2;

    private delegate void WinEventDelegate(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr hwnd,
        int idObject,
        int idChild,
        uint eventThread,
        uint eventTime);

    [DllImport("user32.dll")]
    private static extern IntPtr SetWinEventHook(
        uint eventMin,
        uint eventMax,
        IntPtr hmodWinEventProc,
        WinEventDelegate lpfnWinEventProc,
        uint idProcess,
        uint idThread,
        uint dwFlags);

    [DllImport("user32.dll")]
    private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);
}
