using System.Runtime.InteropServices;
using WinTab.Core.Enums;
using WinTab.Core.Interfaces;

namespace WinTab.Platform.Win32;

/// <summary>
/// Win32 implementation of <see cref="IHotKeyManager"/>.
/// Creates a hidden message-only window (HWND_MESSAGE parent) to receive WM_HOTKEY messages,
/// and translates registered IDs to <see cref="HotKeyAction"/> values.
/// </summary>
public sealed class GlobalHotKeyManager : IHotKeyManager
{
    private IntPtr _hwnd;
    private ushort _classAtom;
    private bool _disposed;

    /// <summary>
    /// Maps registered hot key IDs to <see cref="HotKeyAction"/>.
    /// The hot key ID space is defined by the caller; typically the ordinal of the enum.
    /// </summary>
    private readonly Dictionary<int, HotKeyAction> _registeredKeys = [];

    // The WndProc delegate MUST be stored to prevent GC.
    private readonly NativeMethods.WndProc _wndProc;

    // ─── Events ─────────────────────────────────────────────────────────────

    public event EventHandler<HotKeyAction>? HotKeyPressed;

    // ─── Constructor ────────────────────────────────────────────────────────

    public GlobalHotKeyManager()
    {
        _wndProc = WndProc;
        CreateMessageWindow();
    }

    // ─── IHotKeyManager ─────────────────────────────────────────────────────

    /// <summary>
    /// Register a global hot key.
    /// </summary>
    /// <param name="id">Unique integer ID. Caller should use <c>(int)HotKeyAction.XXX</c>.</param>
    /// <param name="modifiers">Win32 modifier flags (MOD_ALT, MOD_CONTROL, etc.).</param>
    /// <param name="key">Virtual key code.</param>
    /// <returns>True if registration succeeded.</returns>
    public bool Register(int id, uint modifiers, uint key)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (_hwnd == IntPtr.Zero)
            return false;

        if (!NativeMethods.RegisterHotKey(_hwnd, id, modifiers, key))
            return false;

        // Map the integer ID to a HotKeyAction if it falls within the enum range.
        if (Enum.IsDefined(typeof(HotKeyAction), id))
            _registeredKeys[id] = (HotKeyAction)id;
        else
            _registeredKeys[id] = default;

        return true;
    }

    /// <summary>
    /// Unregister a previously registered global hot key.
    /// </summary>
    public bool Unregister(int id)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (_hwnd == IntPtr.Zero)
            return false;

        bool result = NativeMethods.UnregisterHotKey(_hwnd, id);
        _registeredKeys.Remove(id);
        return result;
    }

    /// <summary>
    /// Unregister all previously registered global hot keys.
    /// </summary>
    public void UnregisterAll()
    {
        if (_hwnd == IntPtr.Zero)
            return;

        foreach (int id in _registeredKeys.Keys)
            NativeMethods.UnregisterHotKey(_hwnd, id);

        _registeredKeys.Clear();
    }

    // ─── Message-Only Window ────────────────────────────────────────────────

    private void CreateMessageWindow()
    {
        IntPtr hInstance = NativeMethods.GetModuleHandle(null);
        string className = $"WinTab_HotKey_{Guid.NewGuid():N}";

        var wc = new NativeMethods.WNDCLASS
        {
            lpfnWndProc = _wndProc,
            hInstance = hInstance,
            lpszClassName = className
        };

        _classAtom = NativeMethods.RegisterClass(ref wc);
        if (_classAtom == 0)
            return;

        _hwnd = NativeMethods.CreateWindowEx(
            0, className, string.Empty, 0,
            0, 0, 0, 0,
            NativeConstants.HWND_MESSAGE, // message-only window
            IntPtr.Zero, hInstance, IntPtr.Zero);
    }

    private IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        if (msg == (uint)NativeConstants.WM_HOTKEY)
        {
            int id = wParam.ToInt32();
            if (_registeredKeys.TryGetValue(id, out HotKeyAction action))
                HotKeyPressed?.Invoke(this, action);
        }

        return NativeMethods.DefWindowProc(hWnd, msg, wParam, lParam);
    }

    /// <summary>
    /// Pumps the message queue for the hidden window.
    /// Call this periodically from a timer or message loop if not running on a WPF dispatcher thread.
    /// When used inside a WPF application, the dispatcher message loop handles this automatically.
    /// </summary>
    public void ProcessMessages()
    {
        while (NativeMethods.PeekMessage(out NativeMethods.MSG msg, _hwnd, 0, 0, NativeMethods.PM_REMOVE))
        {
            NativeMethods.TranslateMessage(ref msg);
            NativeMethods.DispatchMessage(ref msg);
        }
    }

    // ─── IDisposable ────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        UnregisterAll();

        if (_hwnd != IntPtr.Zero)
        {
            NativeMethods.DestroyWindow(_hwnd);
            _hwnd = IntPtr.Zero;
        }
    }
}
