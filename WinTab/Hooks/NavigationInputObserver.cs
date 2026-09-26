using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

internal enum NavigationPointerKind { MiddleDown, MiddleUp, OtherDown }
/// <param name="FromWinTab">The input was synthesized by WinTab itself (a double-click close), not by the user or another tool.</param>
/// <param name="DelayMs">How long after the input happened it reached WinTab. The tabs a click starts from are read
/// when it arrives, so a long delay would let Explorer's new tab slip into them.</param>
internal readonly record struct NavigationPointerInput(NavigationPointerKind Kind, Point Point, bool FromWinTab, int DelayMs = 0);

/// <summary>
/// Observes the mouse buttons and the keyboard through Raw Input on a thread of its own. Raw Input only
/// reports input and never sits in its path, so unlike a low-level hook it cannot slow the user's mouse
/// down, has no time limit to exceed, and cannot be removed by Windows without notice: a low-level hook that
/// once answers too late is dropped silently, and every middle click after that is lost without a trace.
/// Input and tab snapshots stay off the WPF thread. Delivery delay is logged because Raw Input is asynchronous;
/// thread priority reduces that delay, but does not guarantee ordering against Explorer's message pump.
/// </summary>
internal sealed class NavigationInputObserver : IDisposable
{
    private const int RegistrationCheckMs = 30_000;
    private static int _classCounter;
    private static readonly object RegistrationGate = new();
    private readonly Action<NavigationPointerInput> _pointer;
    private readonly Action _cancel;
    private readonly Action<nint> _foreground;
    private readonly WindowProcedure _windowProcedure;
    private readonly WinEventDelegate _foregroundCallback;
    private readonly ManualResetEventSlim _started = new();
    /// <summary>Every mouse movement arrives as input; one buffer serves them all on the observer thread.</summary>
    private readonly byte[] _input = new byte[128];
    private readonly Thread _thread;
    private Exception? _startError;
    private uint _threadId;
    private nint _window;
    private volatile bool _stopping;
    private int _disposed;

    public NavigationInputObserver(Action<NavigationPointerInput> pointer, Action cancel, Action<nint> foreground)
    {
        _pointer = pointer;
        _cancel = cancel;
        _foreground = foreground;
        _windowProcedure = OnMessage;
        _foregroundCallback = (_, _, window, _, _, _, _) =>
        {
            if (!_stopping)
            {
                try { _foreground(window); }
                catch (Exception exception) { ExplorerDebugLog.Write($"Navigation foreground observation failed: {exception.Message}"); }
            }
        };
        // Only a few native handles are read per click. Prefer prompt delivery without making the thread
        // do any COM, UIA or disk work while it records the click's starting tabs.
        _thread = new Thread(Run) { IsBackground = true, Name = "WinTab navigation input observer", Priority = ThreadPriority.Highest };
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
        nint foreground = 0;
        string? className = null;
        try
        {
            _threadId = WinApi.GetCurrentThreadId();
            className = "WinTab.NavigationInput." + Interlocked.Increment(ref _classCounter) + "." + Environment.ProcessId;
            var windowClass = new WindowClass
            {
                Size = (uint)Marshal.SizeOf<WindowClass>(),
                Procedure = Marshal.GetFunctionPointerForDelegate(_windowProcedure),
                Instance = GetModuleHandle(null),
                ClassName = className
            };
            if (RegisterClassEx(ref windowClass) == 0)
                throw new Win32Exception(Marshal.GetLastWin32Error());
            _window = CreateWindowEx(0, className, className, 0, 0, 0, 0, 0, MessageOnlyParent, 0, windowClass.Instance, 0);
            if (_window == 0)
                throw new Win32Exception(Marshal.GetLastWin32Error());
            Register();
            foreground = WinApi.SetWinEventHook(WinApi.EVENT_SYSTEM_FOREGROUND, WinApi.EVENT_SYSTEM_FOREGROUND,
                0, _foregroundCallback, 0, 0, WinApi.WINEVENT_OUTOFCONTEXT);
            if (foreground == 0)
                throw new Win32Exception(Marshal.GetLastWin32Error());
            if (SetTimer(_window, 1, RegistrationCheckMs, 0) == 0)
                throw new Win32Exception(Marshal.GetLastWin32Error());
            ExplorerDebugLog.Write($"Navigation input observer started (raw input) thread={_threadId} window={_window}");
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
            if (foreground != 0) WinApi.UnhookWinEvent(foreground);
            if (_window != 0)
            {
                Unregister();
                DestroyWindow(_window);
            }
            if (className != null) UnregisterClass(className, GetModuleHandle(null));
            if (!_stopping)
                ExplorerDebugLog.Write("Navigation input observer stopped unexpectedly");
            _started.Set();
        }
    }

    private void Register()
    {
        lock (RegistrationGate)
        {
            RawInputDevice[] devices =
            [
                new() { UsagePage = 0x01, Usage = MouseUsage, Flags = InputSink, Target = _window },
                new() { UsagePage = 0x01, Usage = KeyboardUsage, Flags = InputSink, Target = _window }
            ];
            if (!RegisterRawInputDevices(devices, (uint)devices.Length, (uint)Marshal.SizeOf<RawInputDevice>()))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
    }

    private void Unregister()
    {
        try
        {
            lock (RegistrationGate)
            {
                // A replacement may own these process-wide usages when an old thread finally exits. Keep
                // checking ownership and removing it atomic with registrations made by replacement observers.
                var devices = ReadRegistrations().Where(device => device.UsagePage == 0x01 && device.Target == _window &&
                        device.Usage is MouseUsage or KeyboardUsage)
                    .Select(device => new RawInputDevice { UsagePage = device.UsagePage, Usage = device.Usage, Flags = Remove }).ToArray();
                if (devices.Length > 0 && !RegisterRawInputDevices(devices, (uint)devices.Length, (uint)Marshal.SizeOf<RawInputDevice>()))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
            }
        }
        catch (Win32Exception exception)
        {
            ExplorerDebugLog.Write($"Navigation raw input cleanup failed: {exception.Message}");
        }
    }

    private static RawInputDevice[] ReadRegistrations()
    {
        uint count = 0;
        var size = (uint)Marshal.SizeOf<RawInputDevice>();
        if (GetRegisteredRawInputDevices(null, ref count, size) == uint.MaxValue)
            throw new Win32Exception(Marshal.GetLastWin32Error());
        if (count == 0)
            return [];
        var devices = new RawInputDevice[count];
        var read = GetRegisteredRawInputDevices(devices, ref count, size);
        if (read == uint.MaxValue)
            throw new Win32Exception(Marshal.GetLastWin32Error());
        return devices[..(int)read];
    }

    /// <summary>
    /// Raw Input has one registration per device type per process. If another part of the process ever took the
    /// mouse or keyboard over, clicks would stop arriving here; the registration is checked and restored.
    /// </summary>
    private void EnsureRegistered()
    {
        try
        {
            var devices = ReadRegistrations();
            var mouse = devices.Any(device => device.UsagePage == 0x01 && device.Usage == MouseUsage && device.Target == _window);
            var keyboard = devices.Any(device => device.UsagePage == 0x01 && device.Usage == KeyboardUsage && device.Target == _window);
            if (mouse && keyboard)
                return;
            ExplorerDebugLog.Write($"Navigation raw input registration was lost mouse={mouse} keyboard={keyboard}; registering again");
            if (!_stopping)
                Register();
        }
        catch (Win32Exception exception)
        {
            ExplorerDebugLog.Write($"Navigation raw input registration check/recovery failed: {exception.Message}");
        }
    }

    private nint OnMessage(nint window, uint message, nint wParam, nint lParam)
    {
        if (message == WmInput && !_stopping)
        {
            try { OnInput(lParam); }
            catch (Exception exception) { ExplorerDebugLog.Write($"Navigation input observation failed: {exception.Message}"); }
        }
        else if (message == WmTimer && !_stopping)
        {
            EnsureRegistered();
            return 0;
        }
        // WM_INPUT must reach DefWindowProc so that Windows can release the input's data.
        return DefWindowProc(window, message, wParam, lParam);
    }

    private void OnInput(nint handle)
    {
        var headerSize = (uint)Marshal.SizeOf<RawInputHeader>();
        var size = (uint)_input.Length;
        if (GetRawInputData(handle, RidInput, _input, ref size, headerSize) == uint.MaxValue || size < headerSize)
            return;
        var data = _input.AsSpan(0, (int)size);
        var header = MemoryMarshal.Read<RawInputHeader>(data);
        var body = data[(int)headerSize..];
        if (header.Type == TypeMouse && body.Length >= Marshal.SizeOf<RawMouse>())
        {
            var mouse = MemoryMarshal.Read<RawMouse>(body);
            var kinds = Classify(mouse.ButtonFlags);
            if (kinds.Count == 0)
                return;
            var delay = unchecked(Environment.TickCount - GetMessageTime());
            var position = GetMessagePos();
            var point = new Point((short)(position & 0xFFFF), (short)((position >> 16) & 0xFFFF));
            // SendInput reports no device. Only input carrying WinTab's own signature is WinTab's.
            var fromWinTab = header.Device == 0 && mouse.ExtraInformation == (uint)MouseSimulator.InjectionSignature;
            foreach (var kind in kinds)
                _pointer(new NavigationPointerInput(kind, point, fromWinTab, delay));
        }
        else if (header.Type == TypeKeyboard && body.Length >= Marshal.SizeOf<RawKeyboard>())
        {
            var keyboard = MemoryMarshal.Read<RawKeyboard>(body);
            if (header.Device != 0 && keyboard.Message is WinApi.WM_KEYDOWN or WinApi.WM_SYSKEYDOWN)
                _cancel();
        }
    }

    /// <summary>
    /// The button transitions of one mouse report that make up a middle click, in the order they are handled:
    /// the middle button's down and up, and another button's down, which starts a different action. Pointer
    /// movement and the wheel are not part of it: Explorer opens the folder's tab even when the pointer or the
    /// wheel moves while the wheel button is pressed.
    /// </summary>
    internal static IReadOnlyList<NavigationPointerKind> Classify(ushort buttonFlags)
    {
        // Most reports are pointer movement; they must not cost an allocation each.
        if ((buttonFlags & (MiddleDown | MiddleUp | OtherButtonsDown)) == 0)
            return [];
        var kinds = new List<NavigationPointerKind>(2);
        if ((buttonFlags & MiddleDown) != 0)
            kinds.Add(NavigationPointerKind.MiddleDown);
        if ((buttonFlags & OtherButtonsDown) != 0)
            kinds.Add(NavigationPointerKind.OtherDown);
        if ((buttonFlags & MiddleUp) != 0)
            kinds.Add(NavigationPointerKind.MiddleUp);
        return kinds;
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

    internal const ushort LeftDown = 0x0001, RightDown = 0x0004, MiddleDown = 0x0010, MiddleUp = 0x0020,
        Button4Down = 0x0040, Button5Down = 0x0100, Wheel = 0x0400, HorizontalWheel = 0x0800;
    private const ushort OtherButtonsDown = LeftDown | RightDown | Button4Down | Button5Down;
    private const ushort MouseUsage = 0x02, KeyboardUsage = 0x06;
    private const uint InputSink = 0x00000100, Remove = 0x00000001;
    private const uint WmInput = 0x00FF, WmTimer = 0x0113, RidInput = 0x10000003, TypeMouse = 0, TypeKeyboard = 1;
    private static readonly nint MessageOnlyParent = -3;

    private delegate nint WindowProcedure(nint window, uint message, nint wParam, nint lParam);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WindowClass
    {
        public uint Size, Style;
        public nint Procedure;
        public int ClassExtra, WindowExtra;
        public nint Instance, Icon, Cursor, Background;
        public string? MenuName;
        public string ClassName;
        public nint SmallIcon;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawInputDevice { public ushort UsagePage, Usage; public uint Flags; public nint Target; }
    [StructLayout(LayoutKind.Sequential)]
    private struct RawInputHeader { public uint Type, Size; public nint Device, WParam; }
    [StructLayout(LayoutKind.Explicit)]
    private struct RawMouse
    {
        [FieldOffset(0)] public ushort Flags;
        [FieldOffset(4)] public ushort ButtonFlags;
        [FieldOffset(6)] public ushort ButtonData;
        [FieldOffset(8)] public uint RawButtons;
        [FieldOffset(12)] public int LastX;
        [FieldOffset(16)] public int LastY;
        [FieldOffset(20)] public uint ExtraInformation;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct RawKeyboard { public ushort MakeCode, Flags, Reserved, VirtualKey; public uint Message, ExtraInformation; }

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern ushort RegisterClassEx(ref WindowClass windowClass);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern bool UnregisterClass(string className, nint instance);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern nint CreateWindowEx(uint exStyle, string className, string windowName, uint style,
        int x, int y, int width, int height, nint parent, nint menu, nint instance, nint parameter);
    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(nint window);
    [DllImport("user32.dll")]
    private static extern nint DefWindowProc(nint window, uint message, nint wParam, nint lParam);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern nint GetModuleHandle(string? moduleName);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterRawInputDevices(RawInputDevice[] devices, uint count, uint size);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetRegisteredRawInputDevices([Out] RawInputDevice[]? devices, ref uint count, uint size);
    [DllImport("user32.dll")]
    private static extern uint GetRawInputData(nint input, uint command, [Out] byte[] data, ref uint size, uint headerSize);
    [DllImport("user32.dll")]
    private static extern uint GetMessagePos();
    [DllImport("user32.dll")]
    private static extern int GetMessageTime();
    [DllImport("user32.dll", SetLastError = true)]
    private static extern nuint SetTimer(nint window, nuint id, uint elapseMs, nint callback);
}
