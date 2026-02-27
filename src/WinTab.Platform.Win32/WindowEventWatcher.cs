using WinTab.Core.Interfaces;

namespace WinTab.Platform.Win32;

/// <summary>
/// Win32 implementation of <see cref="IWindowEventSource"/>.
/// Uses SetWinEventHook to receive system-wide window events and translates
/// them into strongly-typed .NET events.
/// </summary>
public sealed class WindowEventWatcher : IWindowEventSource
{
    private readonly List<IntPtr> _hookHandles = [];
    private bool _disposed;

    // The delegates MUST be stored as fields to prevent garbage collection
    // while the native hooks are active.
    private readonly NativeMethods.WinEventDelegate _showCallback;
    private readonly NativeMethods.WinEventDelegate _destroyCallback;
    private readonly NativeMethods.WinEventDelegate _foregroundCallback;

    // ─── Events ─────────────────────────────────────────────────────────────

    public event EventHandler<IntPtr>? WindowShown;
    public event EventHandler<IntPtr>? WindowDestroyed;
    public event EventHandler<IntPtr>? WindowForegroundChanged;

    // ─── Constructor ────────────────────────────────────────────────────────

    public WindowEventWatcher()
    {
        const uint flags = NativeConstants.WINEVENT_OUTOFCONTEXT |
                           NativeConstants.WINEVENT_SKIPOWNPROCESS;

        // Initialise and store all delegates to prevent GC.
        _showCallback           = OnWindowShow;
        _destroyCallback        = OnWindowDestroy;
        _foregroundCallback     = OnForeground;

        InstallHook(NativeConstants.EVENT_OBJECT_SHOW,            _showCallback, flags);
        InstallHook(NativeConstants.EVENT_OBJECT_DESTROY,         _destroyCallback, flags);
        InstallHook(NativeConstants.EVENT_SYSTEM_FOREGROUND,      _foregroundCallback, flags);
    }

    // ─── Private Helpers ────────────────────────────────────────────────────

    private void InstallHook(uint eventId, NativeMethods.WinEventDelegate callback, uint flags)
    {
        IntPtr hook = NativeMethods.SetWinEventHook(
            eventId, eventId,
            IntPtr.Zero, callback,
            0, 0, flags);

        if (hook != IntPtr.Zero)
            _hookHandles.Add(hook);
    }

    /// <summary>
    /// Returns true when the event should be processed (OBJID_WINDOW and idChild == 0).
    /// Filters out non-window object events.
    /// </summary>
    private static bool ShouldProcess(IntPtr hwnd, int idObject, int idChild)
    {
        return hwnd != IntPtr.Zero &&
               idObject == NativeConstants.OBJID_WINDOW &&
               idChild == 0;
    }

    // ─── Hook Callbacks ─────────────────────────────────────────────────────

    private void OnWindowShow(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (ShouldProcess(hwnd, idObject, idChild))
            WindowShown?.Invoke(this, hwnd);
    }

    private void OnWindowDestroy(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (ShouldProcess(hwnd, idObject, idChild))
            WindowDestroyed?.Invoke(this, hwnd);
    }

    private void OnForeground(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (ShouldProcess(hwnd, idObject, idChild))
            WindowForegroundChanged?.Invoke(this, hwnd);
    }

    // ─── IDisposable ────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        foreach (IntPtr hook in _hookHandles)
        {
            NativeMethods.UnhookWinEvent(hook);
        }
        _hookHandles.Clear();
    }
}
