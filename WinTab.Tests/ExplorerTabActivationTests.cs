using System;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Interop;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.WinAPI;

internal static class ExplorerTabActivationTests
{

    internal static ExplorerWatcher CreateSelectionWatcher(CancellationTokenSource lifetime)
    {
        var watcher = (ExplorerWatcher)RuntimeHelpers.GetUninitializedObject(typeof(ExplorerWatcher));
        typeof(ExplorerWatcher).GetField("_shellLifetime", BindingFlags.Instance | BindingFlags.NonPublic)!
            .SetValue(watcher, lifetime);
        typeof(ExplorerWatcher).GetField("_currentMerge", BindingFlags.Instance | BindingFlags.NonPublic)!
            .SetValue(watcher, new AsyncLocal<MergeOperation?>());
        return watcher;
    }

    internal sealed class ActivationWindow : IDisposable
    {
        private const string TabClassName = "ShellTabWindowClass";
        private static readonly WindowProcedure TabProcedure = DefWindowProc;
        private static readonly nint Module = GetModuleHandle(null);
        private static readonly ushort TabClass = RegisterTabClass();
        private readonly HwndSource _host;
        private readonly nint[] _tabs;

        public ActivationWindow(int tabCount = 2)
        {
            Check.That(TabClass != 0, "The isolated tab window class must be registered.");
            _host = new HwndSource(new HwndSourceParameters("WinTab isolated tab activation")
            {
                // HWND_MESSAGE cannot be displayed, minimized or activated on the desktop.
                ParentWindow = (nint)(-3),
                WindowStyle = 0,
                Width = 10,
                Height = 10
            });
            _tabs = Enumerable.Range(0, tabCount).Select(tabIndex => CreateTab()).ToArray();
            Check.That(_tabs.Length > 0 && _tabs.All(handle => handle != 0), "All isolated tab windows must exist.");
            _host.AddHook(ProcessMessage);
        }

        public nint Handle => _host.Handle;
        public nint FirstTab => _tabs[0];
        public nint TabAt(int index) => _tabs[index];

        /// <summary>Appends a tab the way Explorer does for a new-tab request; it is not reachable by index.</summary>
        public nint AddTab()
        {
            var tab = CreateTab();
            Check.That(tab != 0, "The appended isolated tab must exist.");
            return tab;
        }

        public nint CreateFolderView(int index)
        {
            foreach (var name in new[] { "SHELLDLL_DefView", "DirectUIHWND", "CtrlNotifySink", "DUIViewWndClassName" })
            {
                var windowClass = new NativeWindowClass { Procedure = TabProcedure, Instance = Module, ClassName = name };
                RegisterClass(ref windowClass);
            }
            var host = CreateWindowEx(0, "DUIViewWndClassName", string.Empty, 0x50000000,
                0, 0, 1, 1, _tabs[index], 0, Module, 0);
            var bridge = CreateWindowEx(0, "DirectUIHWND", string.Empty, 0x50000000,
                0, 0, 1, 1, host, 0, Module, 0);
            var sink = CreateWindowEx(0, "CtrlNotifySink", string.Empty, 0x50000000,
                0, 0, 1, 1, bridge, 0, Module, 0);
            var shellView = CreateWindowEx(0, "SHELLDLL_DefView", string.Empty, 0x50000000,
                0, 0, 1, 1, sink, 0, Module, 0);
            return CreateWindowEx(0, "DirectUIHWND", string.Empty, 0x50000000,
                0, 0, 1, 1, shellView, 0, Module, 0);
        }

        public nint ActiveTab => WinApi.FindWindowEx(Handle, 0, TabClassName, null);

        /// <summary>Creates Explorer's frame-level focus host without showing or activating a window.</summary>
        public nint CreateFrameInputSite(string className = "InputSiteWindowClass")
        {
            var windowClass = new NativeWindowClass { Procedure = TabProcedure, Instance = Module, ClassName = className };
            RegisterClass(ref windowClass);
            var site = CreateWindowEx(0, className, string.Empty, 0x50000000,
                0, 0, 1, 1, Handle, 0, Module, 0);
            Check.That(site != 0, "The frame-level input site must exist before testing focus handoff.");
            return site;
        }
        public int CommandsWhileMinimized { get; private set; }
        public int SwitchDelayMs { get; set; }

        /// <summary>Registers the isolated tab class so other test frames can host the same tab windows.</summary>
        internal static void EnsureTabClassRegistered() =>
            Check.That(TabClass != 0, "The isolated tab window class must be registered.");

        public void SetActive(int index) => WinApi.SetWindowPos(_tabs[index], 0, 0, 0, 0, 0, 0x0013);

        private nint CreateTab() => CreateWindowEx(0, TabClassName, string.Empty, 0x50000000,
            0, 0, 1, 1, Handle, 0, Module, 0);

        private nint ProcessMessage(nint handle, int message, nint parameter, nint argument, ref bool handled)
        {
            if ((uint)message != WinApi.WM_COMMAND || parameter != 0xA221)
                return 0;
            handled = true;
            if (WinApi.IsIconic(handle))
            {
                CommandsWhileMinimized++;
                return 0;
            }
            if (SwitchDelayMs > 0)
                Thread.Sleep(SwitchDelayMs);
            var index = (int)argument - 1;
            if (index >= 0 && index < _tabs.Length)
                SetActive(index);
            return 0;
        }

        private static ushort RegisterTabClass()
        {
            var windowClass = new NativeWindowClass
            {
                Procedure = TabProcedure,
                Instance = Module,
                ClassName = TabClassName
            };
            return RegisterClass(ref windowClass);
        }

        public void Dispose() => _host.Dispose();
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

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern nint GetModuleHandle(string? moduleName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern nint CreateWindowEx(uint extendedStyle, string className, string name, uint style,
        int left, int top, int width, int height, nint parent, nint menu, nint instance, nint parameter);
}
