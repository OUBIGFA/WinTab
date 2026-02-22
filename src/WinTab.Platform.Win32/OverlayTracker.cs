using System.Windows.Threading;

namespace WinTab.Platform.Win32;

/// <summary>
/// Tracks a target window's position, size, z-order, and visibility state
/// using a combination of SetWinEventHook and a fallback polling timer.
/// </summary>
public sealed class OverlayTracker : IDisposable
{
    private IntPtr _targetHwnd;
    private bool _disposed;
    private bool _isTracking;

    // WinEvent hooks for the target window's process/thread.
    private readonly List<IntPtr> _hookHandles = [];

    // Prevent GC of the delegates while hooks are active.
    private NativeMethods.WinEventDelegate? _locationCallback;
    private NativeMethods.WinEventDelegate? _minimizeStartCallback;
    private NativeMethods.WinEventDelegate? _minimizeEndCallback;
    private NativeMethods.WinEventDelegate? _foregroundCallback;

    // Fallback polling timer (200 ms).
    private DispatcherTimer? _pollTimer;

    // Cached state for change detection.
    private NativeStructs.RECT _lastRect;
    private bool _lastMinimized;
    private bool _lastActive;

    // ─── Events ─────────────────────────────────────────────────────────────

    /// <summary>Raised when the tracked window's position or size changes.</summary>
    public event Action<int, int, int, int>? PositionChanged;

    /// <summary>Raised when the tracked window is minimized.</summary>
    public event Action? WindowMinimized;

    /// <summary>Raised when the tracked window is restored from minimized state.</summary>
    public event Action? WindowRestored;

    /// <summary>Raised when the tracked window becomes the foreground window.</summary>
    public event Action? WindowActivated;

    /// <summary>Raised when the tracked window loses foreground status.</summary>
    public event Action? WindowDeactivated;

    // ─── Public API ─────────────────────────────────────────────────────────

    /// <summary>
    /// Begin tracking the specified window.
    /// </summary>
    /// <param name="targetHwnd">Handle of the window to track.</param>
    public void Track(IntPtr targetHwnd)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (_isTracking)
            StopTracking();

        _targetHwnd = targetHwnd;
        if (_targetHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_targetHwnd))
            return;

        _isTracking = true;

        // Snapshot initial state.
        RefreshState();

        // Determine target process/thread for scoped hooks.
        uint threadId = NativeMethods.GetWindowThreadProcessId(_targetHwnd, out uint processId);

        const uint flags = NativeConstants.WINEVENT_OUTOFCONTEXT;

        _locationCallback     = OnLocationChange;
        _minimizeStartCallback = OnMinimizeStart;
        _minimizeEndCallback   = OnMinimizeEnd;
        _foregroundCallback    = OnForeground;

        InstallHook(NativeConstants.EVENT_OBJECT_LOCATIONCHANGE, _locationCallback, processId, threadId, flags);
        InstallHook(NativeConstants.EVENT_SYSTEM_MINIMIZESTART,  _minimizeStartCallback, processId, threadId, flags);
        InstallHook(NativeConstants.EVENT_SYSTEM_MINIMIZEEND,    _minimizeEndCallback, processId, threadId, flags);
        InstallHook(NativeConstants.EVENT_SYSTEM_FOREGROUND,     _foregroundCallback, 0, 0, flags);

        // Fallback polling timer — 200 ms.
        _pollTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(200)
        };
        _pollTimer.Tick += OnPollTick;
        _pollTimer.Start();
    }

    /// <summary>
    /// Stop tracking the current window and unhook all event sources.
    /// </summary>
    public void StopTracking()
    {
        _isTracking = false;

        _pollTimer?.Stop();
        _pollTimer = null;

        foreach (IntPtr hook in _hookHandles)
            NativeMethods.UnhookWinEvent(hook);
        _hookHandles.Clear();

        _locationCallback = null;
        _minimizeStartCallback = null;
        _minimizeEndCallback = null;
        _foregroundCallback = null;

        _targetHwnd = IntPtr.Zero;
    }

    // ─── Private Helpers ────────────────────────────────────────────────────

    private void InstallHook(uint eventId, NativeMethods.WinEventDelegate callback,
        uint processId, uint threadId, uint flags)
    {
        IntPtr hook = NativeMethods.SetWinEventHook(
            eventId, eventId,
            IntPtr.Zero, callback,
            processId, threadId, flags);

        if (hook != IntPtr.Zero)
            _hookHandles.Add(hook);
    }

    private void RefreshState()
    {
        if (!NativeMethods.GetWindowRect(_targetHwnd, out _lastRect))
            return;

        var placement = new NativeStructs.WINDOWPLACEMENT
        {
            length = (uint)System.Runtime.InteropServices.Marshal.SizeOf<NativeStructs.WINDOWPLACEMENT>()
        };
        NativeMethods.GetWindowPlacement(_targetHwnd, ref placement);

        _lastMinimized = placement.showCmd == (uint)NativeConstants.SW_SHOWMINIMIZED;
        _lastActive = NativeMethods.GetForegroundWindow() == _targetHwnd;
    }

    // ─── Hook Callbacks ─────────────────────────────────────────────────────

    private void OnLocationChange(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (hwnd != _targetHwnd || idObject != NativeConstants.OBJID_WINDOW || idChild != 0)
            return;

        if (NativeMethods.GetWindowRect(_targetHwnd, out NativeStructs.RECT rect))
        {
            if (!RectEquals(rect, _lastRect))
            {
                _lastRect = rect;
                PositionChanged?.Invoke(rect.Left, rect.Top, rect.Width, rect.Height);
            }
        }
    }

    private void OnMinimizeStart(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (hwnd != _targetHwnd) return;
        _lastMinimized = true;
        WindowMinimized?.Invoke();
    }

    private void OnMinimizeEnd(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (hwnd != _targetHwnd) return;
        _lastMinimized = false;
        WindowRestored?.Invoke();
    }

    private void OnForeground(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (hwnd == _targetHwnd)
        {
            if (!_lastActive)
            {
                _lastActive = true;
                WindowActivated?.Invoke();
            }
        }
        else
        {
            if (_lastActive)
            {
                _lastActive = false;
                WindowDeactivated?.Invoke();
            }
        }
    }

    // ─── Polling Fallback ───────────────────────────────────────────────────

    private void OnPollTick(object? sender, EventArgs e)
    {
        if (!_isTracking || _targetHwnd == IntPtr.Zero)
            return;

        if (!NativeMethods.IsWindow(_targetHwnd))
        {
            StopTracking();
            return;
        }

        // Position / size check.
        if (NativeMethods.GetWindowRect(_targetHwnd, out NativeStructs.RECT rect))
        {
            if (!RectEquals(rect, _lastRect))
            {
                _lastRect = rect;
                PositionChanged?.Invoke(rect.Left, rect.Top, rect.Width, rect.Height);
            }
        }

        // Minimize state check.
        var placement = new NativeStructs.WINDOWPLACEMENT
        {
            length = (uint)System.Runtime.InteropServices.Marshal.SizeOf<NativeStructs.WINDOWPLACEMENT>()
        };
        if (NativeMethods.GetWindowPlacement(_targetHwnd, ref placement))
        {
            bool isMinimized = placement.showCmd == (uint)NativeConstants.SW_SHOWMINIMIZED;
            if (isMinimized != _lastMinimized)
            {
                _lastMinimized = isMinimized;
                if (isMinimized)
                    WindowMinimized?.Invoke();
                else
                    WindowRestored?.Invoke();
            }
        }

        // Active state check.
        bool isActive = NativeMethods.GetForegroundWindow() == _targetHwnd;
        if (isActive != _lastActive)
        {
            _lastActive = isActive;
            if (isActive)
                WindowActivated?.Invoke();
            else
                WindowDeactivated?.Invoke();
        }
    }

    private static bool RectEquals(NativeStructs.RECT a, NativeStructs.RECT b)
    {
        return a.Left == b.Left && a.Top == b.Top && a.Right == b.Right && a.Bottom == b.Bottom;
    }

    // ─── IDisposable ────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        StopTracking();
    }
}
