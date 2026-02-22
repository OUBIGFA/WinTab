using System.Runtime.InteropServices;

namespace WinTab.Platform.Win32;

/// <summary>
/// Low-level mouse hook that detects window title bar drag gestures.
/// Uses WH_MOUSE_LL (hook ID 14) to intercept mouse events system-wide.
/// </summary>
public sealed class DragDetector : IDisposable
{
    private IntPtr _hookHandle;
    private bool _disposed;
    private bool _enabled;

    // MUST be stored as a field to prevent GC while the hook is active.
    private readonly NativeMethods.LowLevelMouseProc _hookCallback;

    // Drag tracking state.
    private bool _isDragging;
    private bool _potentialDrag;
    private IntPtr _sourceWindow;
    private NativeStructs.POINT _mouseDownPoint;

    /// <summary>
    /// Movement threshold in pixels before a mouse-down is considered a drag.
    /// </summary>
    private const int DragThreshold = 8;

    // ─── Events ─────────────────────────────────────────────────────────────

    /// <summary>Raised when a drag gesture begins on a title bar.</summary>
    public event Action<IntPtr, int, int>? DragStarted;

    /// <summary>Raised on each mouse move during a drag gesture.</summary>
    public event Action<int, int>? DragMoved;

    /// <summary>Raised when a drag gesture ends (mouse button released).</summary>
    public event Action<int, int>? DragEnded;

    /// <summary>Raised when a drag gesture is cancelled (e.g., right-click).</summary>
    public event Action? DragCancelled;

    // ─── Constructor ────────────────────────────────────────────────────────

    public DragDetector()
    {
        _hookCallback = HookCallback;
    }

    // ─── Public API ─────────────────────────────────────────────────────────

    /// <summary>
    /// Enable the low-level mouse hook.
    /// </summary>
    public void Enable()
    {
        if (_enabled || _disposed) return;

        IntPtr hModule = NativeMethods.GetModuleHandle(null);
        _hookHandle = NativeMethods.SetWindowsHookEx(
            NativeConstants.WH_MOUSE_LL, _hookCallback, hModule, 0);

        _enabled = _hookHandle != IntPtr.Zero;
    }

    /// <summary>
    /// Disable the low-level mouse hook.
    /// </summary>
    public void Disable()
    {
        if (!_enabled) return;

        if (_isDragging)
            CancelDrag();

        if (_hookHandle != IntPtr.Zero)
        {
            NativeMethods.UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
        }

        _enabled = false;
    }

    /// <summary>
    /// Whether the hook is currently active.
    /// </summary>
    public bool IsEnabled => _enabled;

    /// <summary>
    /// Whether a drag gesture is currently in progress.
    /// </summary>
    public bool IsDragging => _isDragging;

    // ─── Hook Callback ──────────────────────────────────────────────────────

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _enabled)
        {
            int message = wParam.ToInt32();

            switch (message)
            {
                case NativeConstants.WM_LBUTTONDOWN:
                    OnMouseDown(lParam);
                    break;

                case NativeConstants.WM_MOUSEMOVE:
                    OnMouseMove(lParam);
                    break;

                case NativeConstants.WM_LBUTTONUP:
                    OnMouseUp(lParam);
                    break;

                case NativeConstants.WM_RBUTTONDOWN:
                case NativeConstants.WM_RBUTTONUP:
                    if (_isDragging || _potentialDrag)
                        CancelDrag();
                    break;
            }
        }

        return NativeMethods.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
    }

    private void OnMouseDown(IntPtr lParam)
    {
        var hookStruct = Marshal.PtrToStructure<NativeStructs.MSLLHOOKSTRUCT>(lParam);

        // Identify the window under the cursor.
        IntPtr hwnd = NativeMethods.WindowFromPoint(hookStruct.pt);
        if (hwnd == IntPtr.Zero) return;

        // Walk up to the top-level (root owner) window.
        IntPtr rootOwner = NativeMethods.GetAncestor(hwnd, NativeConstants.GA_ROOTOWNER);
        if (rootOwner != IntPtr.Zero)
            hwnd = rootOwner;

        // Check if the mouse is in the title bar region.
        if (!IsInTitleBarRegion(hwnd, hookStruct.pt))
            return;

        _potentialDrag = true;
        _sourceWindow = hwnd;
        _mouseDownPoint = hookStruct.pt;
    }

    private void OnMouseMove(IntPtr lParam)
    {
        var hookStruct = Marshal.PtrToStructure<NativeStructs.MSLLHOOKSTRUCT>(lParam);

        if (_potentialDrag && !_isDragging)
        {
            int dx = hookStruct.pt.X - _mouseDownPoint.X;
            int dy = hookStruct.pt.Y - _mouseDownPoint.Y;

            if (dx * dx + dy * dy >= DragThreshold * DragThreshold)
            {
                _isDragging = true;
                _potentialDrag = false;
                DragStarted?.Invoke(_sourceWindow, _mouseDownPoint.X, _mouseDownPoint.Y);
            }
        }

        if (_isDragging)
        {
            DragMoved?.Invoke(hookStruct.pt.X, hookStruct.pt.Y);
        }
    }

    private void OnMouseUp(IntPtr lParam)
    {
        if (_isDragging)
        {
            var hookStruct = Marshal.PtrToStructure<NativeStructs.MSLLHOOKSTRUCT>(lParam);
            _isDragging = false;
            DragEnded?.Invoke(hookStruct.pt.X, hookStruct.pt.Y);
        }

        _potentialDrag = false;
        _sourceWindow = IntPtr.Zero;
    }

    private void CancelDrag()
    {
        _isDragging = false;
        _potentialDrag = false;
        _sourceWindow = IntPtr.Zero;
        DragCancelled?.Invoke();
    }

    // ─── Title Bar Detection ────────────────────────────────────────────────

    /// <summary>
    /// Determines whether the given screen-space point is within the title bar
    /// region of the specified window.
    /// </summary>
    public static bool IsInTitleBarRegion(IntPtr hWnd, NativeStructs.POINT screenPoint)
    {
        if (!NativeMethods.GetWindowRect(hWnd, out NativeStructs.RECT windowRect))
            return false;

        int captionHeight = NativeMethods.GetSystemMetrics(NativeConstants.SM_CYCAPTION);
        int frameHeight = NativeMethods.GetSystemMetrics(NativeConstants.SM_CYFRAME);
        int titleBarBottom = windowRect.Top + frameHeight + captionHeight;

        return screenPoint.X >= windowRect.Left &&
               screenPoint.X <= windowRect.Right &&
               screenPoint.Y >= windowRect.Top &&
               screenPoint.Y <= titleBarBottom;
    }

    // ─── IDisposable ────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Disable();
    }
}
