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
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("a minimized window is restored before switching to its first tab", () => SelectFirstTabAsync(false));
        yield return ("reusing an already active first tab restores its minimized window", () => SelectFirstTabAsync(true));
    }

    private static async Task SelectFirstTabAsync(bool alreadyActive)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            using var lifetime = new CancellationTokenSource();
            var watcher = CreateSelectionWatcher(lifetime);
            using var fixture = new ActivationWindow();
            fixture.SetActive(alreadyActive ? 0 : 1);
            WinApi.ShowWindow(fixture.Handle, 6);
            Check.That(WinApi.IsIconic(fixture.Handle), "The isolated test window must start minimized.");

            var selected = await watcher.SelectTabByHandle(fixture.Handle, fixture.FirstTab, timeoutMs: 400);

            Check.That(selected, "Reusing the first tab must not fail just because its window was minimized.");
            Check.That(!WinApi.IsIconic(fixture.Handle), "Successful tab reuse must restore the target window.");
            Check.Equal(fixture.FirstTab, fixture.ActiveTab, "The first tab must be active after reuse.");
            Check.Equal(0, fixture.CommandsWhileMinimized, "The window must be restored before a tab-switch command is sent.");
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }

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
                WindowStyle = 0x00CF0000,
                ExtendedWindowStyle = 0x00000080,
                PositionX = -32000,
                PositionY = -32000,
                Width = 10,
                Height = 10
            });
            _tabs = Enumerable.Range(0, tabCount).Select(tabIndex => CreateTab()).ToArray();
            Check.That(_tabs.Length > 0 && _tabs.All(handle => handle != 0), "All isolated tab windows must exist.");
            _host.AddHook(ProcessMessage);
        }

        public nint Handle => _host.Handle;
        public nint FirstTab => _tabs[0];
        public nint ActiveTab => WinApi.FindWindowEx(Handle, 0, TabClassName, null);
        public int CommandsWhileMinimized { get; private set; }

        /// <summary>Registers the isolated tab class so other test frames can host the same tab windows.</summary>
        internal static void EnsureTabClassRegistered() =>
            Check.That(TabClass != 0, "The isolated tab window class must be registered.");

        public void SetActive(int index) => WinApi.SetWindowPos(_tabs[index], 0, 0, 0, 0, 0, 0x0013);

        /// <summary>Shows the off-screen frame the way Explorer shows a real window: visible, not activated.</summary>
        public void Show() => WinApi.ShowWindow(Handle, WinApi.SW_SHOWNOACTIVATE);

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
