using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

internal enum NavigationPointerKind { MiddleDown, MiddleUp, OtherDown }
/// <param name="FromWinTab">The input was synthesized by WinTab itself (a double-click close), not by the user or another tool.</param>
internal readonly record struct NavigationPointerInput(NavigationPointerKind Kind, Point Point, bool FromWinTab);

/// <summary>
/// Observation only: always calls the next hook. Unlike H.Hooks' public events, this tells WinTab's own
/// synthesized clicks apart and guarantees the window's tabs are read before Explorer receives the button-down.
/// </summary>
internal sealed class NavigationInputObserver : IDisposable
{
    private delegate nint HookCallback(int code, nint message, nint data);
    private readonly Action<NavigationPointerInput> _pointer;
    private readonly Action _cancel;
    private readonly Action<nint> _foreground;
    private readonly HookCallback _mouseCallback;
    private readonly HookCallback _keyboardCallback;
    private readonly WinEventDelegate _foregroundCallback;
    private readonly ManualResetEventSlim _started = new();
    private readonly Thread _thread;
    private Exception? _startError;
    private uint _threadId;
    private volatile bool _stopping;
    private int _disposed;

    public NavigationInputObserver(Action<NavigationPointerInput> pointer, Action cancel, Action<nint> foreground)
    {
        _pointer = pointer;
        _cancel = cancel;
        _foreground = foreground;
        _mouseCallback = OnMouse;
        _keyboardCallback = OnKeyboard;
        _foregroundCallback = (_, _, window, _, _, _, _) =>
        {
            if (!_stopping)
            {
                try { _foreground(window); }
                catch (Exception exception) { ExplorerDebugLog.Write($"Navigation foreground observation failed: {exception.Message}"); }
            }
        };
        _thread = new Thread(Run) { IsBackground = true, Name = "WinTab navigation input observer" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        if (!_started.Wait(2_000))
        {
            Dispose();
            throw new TimeoutException("Navigation input observer did not start.");
        }
        if (_startError != null)
        {
            Dispose();
            throw new InvalidOperationException("Navigation input observer could not be installed.", _startError);
        }
    }

    private void Run()
    {
        nint mouse = 0, keyboard = 0, foreground = 0;
        try
        {
            _threadId = WinApi.GetCurrentThreadId();
            WinApi.PeekMessage(out _, 0, 0, 0, WinApi.PM_NOREMOVE);
            mouse = SetWindowsHookEx(WinApi.WH_MOUSE_LL, _mouseCallback, 0, 0);
            keyboard = SetWindowsHookEx(WinApi.WH_KEYBOARD_LL, _keyboardCallback, 0, 0);
            foreground = WinApi.SetWinEventHook(WinApi.EVENT_SYSTEM_FOREGROUND, WinApi.EVENT_SYSTEM_FOREGROUND,
                0, _foregroundCallback, 0, 0, WinApi.WINEVENT_OUTOFCONTEXT);
            if (mouse == 0 || keyboard == 0 || foreground == 0)
                throw new Win32Exception(Marshal.GetLastWin32Error());
            _started.Set();
            while (!_stopping && WinApi.GetMessage(out var message, 0, 0, 0) > 0)
            {
                WinApi.TranslateMessage(ref message);
                WinApi.DispatchMessage(ref message);
            }
        }
        catch (Exception exception)
        {
            _startError = exception;
            ExplorerDebugLog.Write($"Navigation input observer failed: {exception.Message}");
        }
        finally
        {
            if (mouse != 0) UnhookWindowsHookEx(mouse);
            if (keyboard != 0) UnhookWindowsHookEx(keyboard);
            if (foreground != 0) WinApi.UnhookWinEvent(foreground);
            _started.Set();
        }
    }

    /// <summary>
    /// The mouse input that makes up a middle click: its button-down and button-up, and another button, which
    /// starts a different action. Pointer movement and the wheel are not observed: Explorer opens the folder's
    /// tab even when the pointer or the wheel moves while the wheel button is pressed.
    /// </summary>
    internal static NavigationPointerKind? Classify(uint message) => message switch
    {
        WinApi.WM_MBUTTONDOWN => NavigationPointerKind.MiddleDown,
        WinApi.WM_MBUTTONUP => NavigationPointerKind.MiddleUp,
        WinApi.WM_LBUTTONDOWN or WinApi.WM_RBUTTONDOWN or WinApi.WM_XBUTTONDOWN => NavigationPointerKind.OtherDown,
        _ => null
    };

    private nint OnMouse(int code, nint message, nint data)
    {
        if (code >= 0 && !_stopping && Classify((uint)message) is { } kind)
        {
            var input = Marshal.PtrToStructure<MouseInput>(data);
            var fromWinTab = (input.Flags & WinApi.LLMHF_INJECTED) != 0 && input.ExtraInfo == MouseSimulator.InjectionSignature;
            try { _pointer(new NavigationPointerInput(kind, input.Point, fromWinTab)); }
            catch (Exception exception) { ExplorerDebugLog.Write($"Navigation pointer observation failed: {exception.Message}"); }
        }
        return CallNextHookEx(0, code, message, data);
    }

    private nint OnKeyboard(int code, nint message, nint data)
    {
        if (code >= 0 && !_stopping && ((uint)message is WinApi.WM_KEYDOWN or WinApi.WM_SYSKEYDOWN))
        {
            var input = Marshal.PtrToStructure<KeyboardInput>(data);
            if ((input.Flags & WinApi.LLKHF_INJECTED) == 0)
            {
                try { _cancel(); }
                catch (Exception exception) { ExplorerDebugLog.Write($"Navigation keyboard observation failed: {exception.Message}"); }
            }
        }
        return CallNextHookEx(0, code, message, data);
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        _stopping = true;
        if (_threadId != 0)
            WinApi.PostThreadMessage(_threadId, WinApi.WM_QUIT, 0, 0);
        if (_thread.Join(2_000))
            _started.Dispose();
        else
            ExplorerDebugLog.Write("Navigation input observer is still stopping.");
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput { public Point Point; public uint Data, Flags, Time; public nuint ExtraInfo; }
    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardInput { public uint Key, ScanCode, Flags, Time; public nuint ExtraInfo; }
    [DllImport("user32.dll", SetLastError = true)]
    private static extern nint SetWindowsHookEx(int kind, HookCallback callback, nint module, uint threadId);
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(nint hook);
    [DllImport("user32.dll")]
    private static extern nint CallNextHookEx(nint hook, int code, nint message, nint data);
}
