using System;
using System.Runtime.InteropServices;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;

namespace WinTab.App.ExplorerTabUtilityPort;

/// <summary>
/// Encapsulates the WinEvent hook lifecycle and COM initialization requirements
/// for detecting Explorer tab creation and registration events.
/// </summary>
public sealed class WinEventHookManager : IDisposable
{
    private readonly Logger _logger;
    private readonly NativeMethods.WinEventDelegate _createEventCallback;
    private IntPtr _createEventHook;
    private bool _disposed;

    public WinEventHookManager(Logger logger, NativeMethods.WinEventDelegate createEventCallback)
    {
        _logger = logger;
        _createEventCallback = createEventCallback ?? throw new ArgumentNullException(nameof(createEventCallback));
    }

    public void Start()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(WinEventHookManager));

        const uint flags = NativeConstants.WINEVENT_OUTOFCONTEXT | NativeConstants.WINEVENT_SKIPOWNPROCESS;
        _createEventHook = NativeMethods.SetWinEventHook(
            NativeConstants.EVENT_OBJECT_CREATE,
            NativeConstants.EVENT_OBJECT_CREATE,
            IntPtr.Zero,
            _createEventCallback,
            0,
            0,
            flags);

        if (_createEventHook == IntPtr.Zero)
        {
            _logger.Warn("WinEventHookManager: EVENT_OBJECT_CREATE hook unavailable.");
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_createEventHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWinEvent(_createEventHook);
            _createEventHook = IntPtr.Zero;
        }
    }
}
