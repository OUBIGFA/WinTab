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
    private readonly NativeMethods.WinEventDelegate _locationChangeCallback;
    private readonly NativeMethods.WinEventDelegate _minimizeStartCallback;
    private readonly NativeMethods.WinEventDelegate _minimizeEndCallback;
    private readonly NativeMethods.WinEventDelegate _foregroundCallback;
    private readonly NativeMethods.WinEventDelegate _moveSizeStartCallback;
    private readonly NativeMethods.WinEventDelegate _moveSizeEndCallback;

    // ─── Events ─────────────────────────────────────────────────────────────

    public event EventHandler<IntPtr>? WindowShown;
    public event EventHandler<IntPtr>? WindowDestroyed;
    public event EventHandler<IntPtr>? WindowLocationChanged;
    public event EventHandler<IntPtr>? WindowMinimized;
    public event EventHandler<IntPtr>? WindowRestored;
    public event EventHandler<IntPtr>? WindowForegroundChanged;
    public event EventHandler<IntPtr>? WindowMoveSizeStarted;
    public event EventHandler<IntPtr>? WindowMoveSizeEnded;

    // ─── Constructor ────────────────────────────────────────────────────────

    public WindowEventWatcher()
    {
        const uint flags = NativeConstants.WINEVENT_OUTOFCONTEXT |
                           NativeConstants.WINEVENT_SKIPOWNPROCESS;

        // Initialise and store all delegates to prevent GC.
        _showCallback           = OnWindowShow;
        _destroyCallback        = OnWindowDestroy;
        _locationChangeCallback = OnWindowLocationChange;
        _minimizeStartCallback  = OnMinimizeStart;
        _minimizeEndCallback    = OnMinimizeEnd;
        _foregroundCallback     = OnForeground;
        _moveSizeStartCallback  = OnMoveSizeStart;
        _moveSizeEndCallback    = OnMoveSizeEnd;

        InstallHook(NativeConstants.EVENT_OBJECT_SHOW,            _showCallback, flags);
        InstallHook(NativeConstants.EVENT_OBJECT_DESTROY,         _destroyCallback, flags);
        InstallHook(NativeConstants.EVENT_OBJECT_LOCATIONCHANGE,  _locationChangeCallback, flags);
        InstallHook(NativeConstants.EVENT_SYSTEM_MINIMIZESTART,   _minimizeStartCallback, flags);
        InstallHook(NativeConstants.EVENT_SYSTEM_MINIMIZEEND,     _minimizeEndCallback, flags);
        InstallHook(NativeConstants.EVENT_SYSTEM_FOREGROUND,      _foregroundCallback, flags);
        InstallHook(NativeConstants.EVENT_SYSTEM_MOVESIZESTART,   _moveSizeStartCallback, flags);
        InstallHook(NativeConstants.EVENT_SYSTEM_MOVESIZEEND,     _moveSizeEndCallback, flags);
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

    private void OnWindowLocationChange(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (ShouldProcess(hwnd, idObject, idChild))
            WindowLocationChanged?.Invoke(this, hwnd);
    }

    private void OnMinimizeStart(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (ShouldProcess(hwnd, idObject, idChild))
            WindowMinimized?.Invoke(this, hwnd);
    }

    private void OnMinimizeEnd(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (ShouldProcess(hwnd, idObject, idChild))
            WindowRestored?.Invoke(this, hwnd);
    }

    private void OnForeground(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (ShouldProcess(hwnd, idObject, idChild))
            WindowForegroundChanged?.Invoke(this, hwnd);
    }

    private void OnMoveSizeStart(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (ShouldProcess(hwnd, idObject, idChild))
            WindowMoveSizeStarted?.Invoke(this, hwnd);
    }

    private void OnMoveSizeEnd(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (ShouldProcess(hwnd, idObject, idChild))
            WindowMoveSizeEnded?.Invoke(this, hwnd);
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
