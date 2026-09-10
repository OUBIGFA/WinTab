using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading;
using WinTab.WinAPI;

/// <summary>
/// An Explorer-like frame that lives on its own message-pumping thread, so window commands cross a
/// thread boundary exactly as they do with the real explorer.exe. The frame can acknowledge close and
/// tab-switch commands slowly, or ignore the close request entirely, to model a busy Explorer.
/// </summary>
internal sealed class RemoteExplorerFrame : IDisposable
{
    private const string OwnClassName = "WinTabTestRemoteFrame";
    private const string ExplorerClassName = "CabinetWClass";
    private const uint WmDestroy = 0x0002;
    private static readonly WindowProcedure FrameProcedure = HandleMessage;
    private static readonly ConcurrentDictionary<nint, RemoteExplorerFrame> Frames = new();
    private static readonly ConcurrentDictionary<string, bool> RegisteredClasses = new();
    private static readonly nint Module = GetModuleHandle(null);
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new();
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

    /// <summary>How long the frame keeps a close request pending before destroying itself.</summary>
    public int CloseDelayMs { get => _closeDelayMs; set => _closeDelayMs = value; }

    /// <summary>How long the frame takes to activate a tab after receiving the switch command.</summary>
    public int SwitchDelayMs { get => _switchDelayMs; set => _switchDelayMs = value; }

    /// <summary>When set, the frame acknowledges the close request but stays open.</summary>
    public bool IgnoreClose { get => _ignoreClose; set => _ignoreClose = value; }

    public void SetActive(int index) => WinApi.SetWindowPos(_tabs[index], 0, 0, 0, 0, 0, 0x0013);

    private void Run()
    {
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
        if (message == WinApi.WM_CLOSE && frame != null)
        {
            if (frame._ignoreClose)
                return 0;
            if (frame._closeDelayMs > 0)
                Thread.Sleep(frame._closeDelayMs);
            DestroyWindow(handle);
            return 0;
        }

        if (message == WinApi.WM_COMMAND && parameter == 0xA221 && frame != null)
        {
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
        _thread.Join(3_000);
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

    [DllImport("user32.dll")]
    private static extern void PostQuitMessage(int exitCode);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern nint GetModuleHandle(string? moduleName);
}
