using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>Explorer-only shortcuts; no RegisterHotKey reservation that would steal browser shortcuts.</summary>
internal sealed class ExplorerSessionShortcutHook : IDisposable
{
    private readonly Dispatcher _dispatcher = Dispatcher.CurrentDispatcher;
    private readonly HookProc _callback;
    private readonly ExplorerShortcutDispatch _dispatch = new();
    private readonly Action<SessionAction> _execute;
    private nint _hook;
    private bool _disposed;
    public event Action<string>? Failed;

    public ExplorerSessionShortcutHook(Action<SessionAction> execute)
    {
        _execute = execute;
        _callback = Callback;
    }

    public void Configure(ExplorerShortcut? group, ExplorerShortcut? tab)
    {
        _dispatcher.VerifyAccess();
        ObjectDisposedException.ThrowIf(_disposed, this);
        _dispatch.Group = group;
        _dispatch.Tab = tab;
        if (group == null && tab == null)
        {
            Unhook();
            return;
        }
        if (_hook != 0) return;
        _hook = SetWindowsHookEx(13, _callback, GetModuleHandle(null), 0);
        if (_hook == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
    }

    private nint Callback(int code, nint message, nint data)
    {
        if (code < 0 || _disposed) return CallNextHookEx(_hook, code, message, data);
        try
        {
            var down = message == 0x100 || message == 0x104;
            var up = message == 0x101 || message == 0x105;
            if (!down && !up) return CallNextHookEx(_hook, code, message, data);
            var key = Marshal.PtrToStructure<KeyboardData>(data);
            var foreground = WinApi.GetForegroundWindow();
            var explorer = ExplorerWindowDiscovery.IsFileExplorerWindow(foreground);
            var modifiers = (Down(0x11) ? ShortcutModifiers.Control : 0) |
                (Down(0x10) ? ShortcutModifiers.Shift : 0) | (Down(0x12) ? ShortcutModifiers.Alt : 0);
            // Win-modified keys belong to Windows, not these shortcuts.
            explorer &= !Down(0x5B) && !Down(0x5C);
            var consumed = _dispatch.Handle((int)key.Key, down, modifiers, explorer, (key.Flags & 0x12) != 0, out var action);
            if (action is { } command)
            {
                var identity = WindowIdentity.Capture(foreground);
                _dispatcher.BeginInvoke(() =>
                {
                    if (!_disposed && identity.IsCurrent && WinApi.GetForegroundWindow() == identity.Handle)
                        _execute(command);
                }, DispatcherPriority.Background);
            }
            if (consumed) return 1;
        }
        catch (Exception exception)
        {
            // Never unwind through the native callback. Report asynchronously, not on the input path.
            _dispatcher.BeginInvoke(() => Failed?.Invoke(exception.Message));
        }
        return CallNextHookEx(_hook, code, message, data);
    }

    private static bool Down(int key) => (WinApi.GetAsyncKeyState(key) & 0x8000) != 0;
    private void Unhook()
    {
        if (_hook == 0) return;
        if (!UnhookWindowsHookEx(_hook)) throw new Win32Exception(Marshal.GetLastWin32Error());
        _hook = 0;
        _dispatch.Reset();
    }
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Unhook();
        GC.KeepAlive(_callback);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardData { public uint Key, ScanCode, Flags, Time; public nuint ExtraInfo; }
    private delegate nint HookProc(int code, nint message, nint data);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern nint SetWindowsHookEx(int id, HookProc callback, nint module, uint thread);
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(nint hook);
    [DllImport("user32.dll")]
    private static extern nint CallNextHookEx(nint hook, int code, nint message, nint data);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern nint GetModuleHandle(string? name);
}
