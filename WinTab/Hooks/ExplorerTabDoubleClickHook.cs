using System;
using System.Drawing;
using System.Threading;
using H.Hooks;
using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

public sealed class ExplorerTabDoubleClickHook : IHook
{
    private const int SM_CXDOUBLECLK = 36;
    private const int SM_CYDOUBLECLK = 37;
    private const uint GA_ROOT = 2;

    private readonly ExplorerWatcher _explorerWatcher;
    private readonly LowLevelMouseHook _lowLevelMouseHook;
    private readonly ExplorerTabDoubleClickCloseController _controller;
    private readonly Func<bool> _isEnabled;

    public ExplorerTabDoubleClickHook(ExplorerWatcher explorerWatcher, Func<bool>? isEnabled = null)
    {
        _explorerWatcher = explorerWatcher;
        _isEnabled = isEnabled ?? (() => true);
        _controller = new ExplorerTabDoubleClickCloseController(new HookEnvironment(explorerWatcher, _isEnabled));
        _lowLevelMouseHook = new LowLevelMouseHook
        {
            AddKeyboardKeys = true,
            Handling = true
        };
        _lowLevelMouseHook.Down += OnMouseDown;
        _lowLevelMouseHook.Up += OnMouseUp;
    }

    public event Action<string>? StatusChanged;
    public bool IsHookActive => _lowLevelMouseHook.IsStarted;

    public void StartHook() => _lowLevelMouseHook.Start();
    public void StopHook() => _lowLevelMouseHook.Stop();

    private void OnMouseDown(object? sender, MouseEventArgs e)
    {
        if (!IsLeftMouse(e))
            return;

        var decision = _controller.HandleLeftMouseDown(e.Position, Environment.TickCount64);
        if (decision.Handled)
            e.IsHandled = true;
    }

    private void OnMouseUp(object? sender, MouseEventArgs e)
    {
        if (!IsLeftMouse(e))
            return;

        var decision = _controller.HandleLeftMouseUp(Environment.TickCount64);
        if (!decision.Handled)
            return;

        e.IsHandled = true;

        if (decision.CloseRequest is not { } closeRequest)
            return;

        // Hand off to the threadpool so the hook thread is never blocked. A tiny delay lets the suppressed
        // mouse-up message drain before we synthesize the middle-click; without it, on slow systems the
        // injected click can race with the original left-up sequence.
        ThreadPool.QueueUserWorkItem(_ =>
        {
            Thread.Sleep(10);
            try
            {
                MouseSimulator.SendMiddleClick(closeRequest.Point);
                StatusChanged?.Invoke("Closed Explorer tab via middle-click.");
            }
            catch
            {
                // 忽略关闭请求失败，避免后台任务影响钩子线程。
            }
            finally
            {
                try
                {
                    _explorerWatcher.RefreshTabStripBounds(closeRequest.ExplorerWindow);
                }
                catch
                {
                    // 忽略刷新失败，后续鼠标事件会重新计算。
                }
            }
        });
    }

    private static bool IsLeftMouse(MouseEventArgs e) => e.CurrentKey is Key.MouseLeft or Key.LButton;

    private static nint GetExplorerWindowForPoint(Point point)
    {
        var hit = WinApi.WindowFromPoint(point);
        if (hit != 0)
        {
            var root = WinApi.GetAncestor(hit, GA_ROOT);
            if (ExplorerWindowDiscovery.IsFileExplorerWindow(root))
                return root;
        }

        return ExplorerWindowDiscovery.IsFileExplorerForeground(out var foreground) && foreground != 0 ? foreground : 0;
    }

    private sealed class HookEnvironment(ExplorerWatcher explorerWatcher, Func<bool> isEnabled) : IExplorerTabDoubleClickEnvironment
    {
        public bool IsEnabled => isEnabled();
        public int DoubleClickTimeMs => (int)WinApi.GetDoubleClickTime();
        public int DoubleClickWidth => WinApi.GetSystemMetrics(SM_CXDOUBLECLK);
        public int DoubleClickHeight => WinApi.GetSystemMetrics(SM_CYDOUBLECLK);
        public nint ResolveExplorerWindow(Point point) => GetExplorerWindowForPoint(point);
        public bool IsExplorerWindow(nint explorerWindow) => ExplorerWindowDiscovery.IsFileExplorerWindow(explorerWindow);
        public bool IsPointOnTabStrip(Point point, nint explorerWindow) =>
            explorerWatcher.IsPointOnTabStrip(point, explorerWindow);
    }

    public void Dispose()
    {
        StopHook();
        _lowLevelMouseHook.Dispose();
    }
}
