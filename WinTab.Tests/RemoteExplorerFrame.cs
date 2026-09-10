using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.WinAPI;

/// <summary>
/// An Explorer-like frame that lives on its own message-pumping thread, so window commands cross a
/// thread boundary exactly as they do with the real explorer.exe. The frame can acknowledge close and
/// tab-switch commands slowly, ignore the close request entirely, or stop pumping messages for a while
/// to model a busy Explorer.
/// </summary>
internal sealed class RemoteExplorerFrame : IDisposable
{
    private const string OwnClassName = "WinTabTestRemoteFrame";
    private const string ExplorerClassName = "CabinetWClass";
    private const uint WmDestroy = 0x0002;
    private const uint WmBlock = 0x8001;
    private static readonly WindowProcedure FrameProcedure = HandleMessage;
    private static readonly ConcurrentDictionary<nint, RemoteExplorerFrame> Frames = new();
    private static readonly ConcurrentDictionary<string, bool> RegisteredClasses = new();
    private static readonly nint Module = GetModuleHandle(null);
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new();
    private readonly TaskCompletionSource _blocked = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly ConcurrentQueue<string> _trace = new();
    private readonly long _createdAt = Stopwatch.GetTimestamp();
    private readonly int _tabCount;
    private readonly bool _visible;
    private readonly string _className;
    private nint[] _tabs = [];
    private volatile int _closeDelayMs;
    private volatile int _switchDelayMs;
    private volatile bool _ignoreClose;

    /// <param name="visible">Whether the frame starts shown, like a user window, or hidden, like a preloaded frame.</param>
    /// <param name="tabCount">Number of ShellTabWindowClass children.</param>
    /// <param name="explorerClass">Register the frame with Explorer's own window class name so class checks treat it as Explorer.</param>
    public RemoteExplorerFrame(bool visible, int tabCount = 1, bool explorerClass = false)
    {
        ExplorerTabActivationTests.ActivationWindow.EnsureTabClassRegistered();
        _visible = visible;
        _tabCount = tabCount;
        _className = explorerClass ? ExplorerClassName : OwnClassName;
        _thread = new Thread(Run) { IsBackground = true, Name = "WinTab remote test frame" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        Check.That(_ready.Wait(5_000) && Handle != 0 && _tabs.Length == tabCount, "The remote test frame must start with its tabs.");
    }

    public nint Handle { get; private set; }
    public nint Tab => _tabs[0];
    public nint ActiveTab => WinApi.FindWindowEx(Handle, 0, "ShellTabWindowClass", null);
    public bool IsAlive => Handle != 0 && IsWindow(Handle);
    public int CloseCommandCount { get; private set; }
    /// <summary>Whether the frame was on screen at the moment it finished handling a close request.</summary>
    public bool RevealedWhileClosing { get; private set; }
    public ConcurrentQueue<int> SwitchCommands { get; } = new();

    /// <summary>Every message the frame window handled, with arrival time and handling duration, for failure diagnostics.</summary>
    public string Trace => string.Join(" ", _trace.TakeLast(40));

    /// <summary>
    /// Stops the frame's thread from processing messages for <paramref name="durationMs"/>. The thread sleeps
    /// instead of waiting on an event: a managed wait on an STA thread still answers sent messages, which
    /// would defeat the point of modelling a busy Explorer.
    /// </summary>
    public async Task BlockMessagesAsync(int durationMs)
    {
        Check.That(WinApi.PostMessage(Handle, WmBlock, durationMs, 0), "The isolated frame must receive the block request.");
        await _blocked.Task.WaitAsync(TimeSpan.FromSeconds(3));
    }

    /// <summary>How long the frame keeps a close request pending before answering it.</summary>
    public int CloseDelayMs { get => _closeDelayMs; set => _closeDelayMs = value; }

    /// <summary>How long the frame takes to activate a tab after receiving the switch command.</summary>
    public int SwitchDelayMs { get => _switchDelayMs; set => _switchDelayMs = value; }

    /// <summary>When set, the frame acknowledges the close request but stays open.</summary>
    public bool IgnoreClose { get => _ignoreClose; set => _ignoreClose = value; }

    public void SetActive(int index) => WinApi.SetWindowPos(_tabs[index], 0, 0, 0, 0, 0, 0x0013);

    private void Run()
    {
        // The isolated test desktop has no text services. Without this, the first activation of the frame
        // blocks its thread for about two seconds while the IME context is set up, which discards commands
        // sent to it in the meantime and has nothing to do with Explorer.
        ImmDisableIME(0);
        RegisterFrameClass(_className);
        Handle = CreateWindowEx(0x00000080, _className, "WinTab remote test frame", 0x00CF0000,
            -32000, -32000, 10, 10, 0, 0, Module, 0);
        Frames[Handle] = this;
        var tabs = new nint[_tabCount];
        for (var index = 0; index < tabs.Length; index++)
            tabs[index] = CreateWindowEx(0, "ShellTabWindowClass", string.Empty, 0x50000000, 0, 0, 1, 1, Handle, 0, Module, 0);
        _tabs = tabs;
        if (_visible)
            WinApi.ShowWindow(Handle, WinApi.SW_SHOWNOACTIVATE);
        _ready.Set();
        while (WinApi.GetMessage(out var message, 0, 0, 0) > 0)
        {
            WinApi.TranslateMessage(ref message);
            WinApi.DispatchMessage(ref message);
        }
    }

    private static nint HandleMessage(nint handle, uint message, nint parameter, nint argument)
    {
        Frames.TryGetValue(handle, out var frame);
        if (frame == null)
            return HandleFrameMessage(null, handle, message, parameter, argument);
        var arrivedAt = Stopwatch.GetTimestamp();
        try
        {
            return HandleFrameMessage(frame, handle, message, parameter, argument);
        }
        finally
        {
            frame._trace.Enqueue($"{message:X}@{Stopwatch.GetElapsedTime(frame._createdAt, arrivedAt).TotalMilliseconds:F0}+{Stopwatch.GetElapsedTime(arrivedAt).TotalMilliseconds:F0}");
        }
    }

    private static nint HandleFrameMessage(RemoteExplorerFrame? frame, nint handle, uint message, nint parameter, nint argument)
    {
        if (message == WmBlock && frame != null)
        {
            frame._blocked.TrySetResult();
            Thread.Sleep((int)parameter);
            return 0;
        }
        if (message == WinApi.WM_CLOSE && frame != null)
        {
            frame.CloseCommandCount++;
            if (frame._closeDelayMs > 0)
                Thread.Sleep(frame._closeDelayMs);
            frame.RevealedWhileClosing = IsRevealed(handle);
            if (frame._ignoreClose)
                return 0;
            DestroyWindow(handle);
            return 0;
        }

        if (message == WinApi.WM_COMMAND && parameter == 0xA221 && frame != null)
        {
            frame.SwitchCommands.Enqueue((int)argument - 1);
            if (frame._switchDelayMs > 0)
                Thread.Sleep(frame._switchDelayMs);
            var index = (int)argument - 1;
            if (index >= 0 && index < frame._tabs.Length)
                frame.SetActive(index);
            return 0;
        }

        if (message == WmDestroy)
        {
            Frames.TryRemove(handle, out _);
            PostQuitMessage(0);
            return 0;
        }

        return DefWindowProc(handle, message, parameter, argument);
    }

    /// <summary>A window is on screen unless it is a layered window whose opacity has been set to zero.</summary>
    private static bool IsRevealed(nint handle)
    {
        if (!WinApi.IsWindowVisible(handle))
            return false;
        if ((WinApi.GetWindowLong(handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0)
            return true;
        return !WinApi.GetLayeredWindowAttributes(handle, out _, out var alpha, out var flags) ||
            (flags & WinApi.LWA_ALPHA) == 0 || alpha != 0;
    }

    private static void RegisterFrameClass(string className)
    {
        if (!RegisteredClasses.TryAdd(className, true))
            return;
        var windowClass = new NativeWindowClass
        {
            Procedure = FrameProcedure,
            Instance = Module,
            ClassName = className
        };
        Check.That(RegisterClass(ref windowClass) != 0, "The remote frame window class must be registered.");
    }

    public void Dispose()
    {
        if (IsAlive)
        {
            _ignoreClose = false;
            _closeDelayMs = 0;
            WinApi.PostMessage(Handle, WinApi.WM_CLOSE, 0, 0);
        }
        // A frame may still be sleeping through a block or a slow close; give it time to finish and exit.
        Check.That(_thread.Join(5_000), "The isolated frame thread must stop before its wait handle is disposed.");
        _ready.Dispose();
    }

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate nint WindowProcedure(nint handle, uint message, nint parameter, nint argument);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeWindowClass
    {
        public uint Style;
        public WindowProcedure Procedure;
        public int ClassExtraBytes;
        public int WindowExtraBytes;
        public nint Instance;
        public nint Icon;
        public nint Cursor;
        public nint Background;
        public string? MenuName;
        public string ClassName;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern ushort RegisterClass(ref NativeWindowClass windowClass);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint DefWindowProc(nint handle, uint message, nint parameter, nint argument);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern nint CreateWindowEx(uint extendedStyle, string className, string name, uint style,
        int left, int top, int width, int height, nint parent, nint menu, nint instance, nint parameter);

    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(nint handle);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(nint handle);

    [DllImport("imm32.dll")]
    private static extern bool ImmDisableIME(uint threadId);

    [DllImport("user32.dll")]
    private static extern void PostQuitMessage(int exitCode);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern nint GetModuleHandle(string? moduleName);
}
