using System.Runtime.InteropServices;
using System.Text;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;

namespace WinTab.App.ExplorerTabUtilityPort;

/// <summary>
/// Handles optional Explorer tab close behavior: double left-click closes tab.
/// </summary>
public sealed class ExplorerTabMouseHookService : IDisposable
{
    private const string ExplorerWindowClass = "CabinetWClass";
    private const string ExplorerTabClass = "ShellTabWindowClass";

    // The WinUI 3 bridge window that hosts actual tab items in Windows 11 Explorer.
    // Clicks on this class are over real UI content (tab titles, buttons, etc.).
    private const string WinUiBridgeClass = "Microsoft.UI.Content.DesktopChildSiteBridge";

    // The scaffolding window that covers the empty/draggable part of the title bar.
    // Clicks on this class land on blank space and should preserve native double-click
    // behaviour (maximize / restore).
    private const string TitleBarScaffoldingClass = "TITLE_BAR_SCAFFOLDING_WINDOW_CLASS";

    private readonly object _sync = new();
    private readonly Logger _logger;
    private readonly NativeMethods.LowLevelMouseProc _mouseProc;
    private readonly int _doubleClickTimeMs;
    private readonly int _doubleClickWidth;
    private readonly int _doubleClickHeight;

    private IntPtr _hookHandle;
    private bool _enabled;
    private bool _disposed;

    private long _lastClickTick;
    private IntPtr _lastClickTopLevelHandle;
    private NativeStructs.POINT _lastClickPoint;

    public ExplorerTabMouseHookService(AppSettings settings, Logger logger)
    {
        _logger = logger;
        _mouseProc = MouseHookProc;

        _doubleClickTimeMs = (int)NativeMethods.GetDoubleClickTime();
        if (_doubleClickTimeMs <= 0)
            _doubleClickTimeMs = 500;

        _doubleClickWidth = Math.Max(1, NativeMethods.GetSystemMetrics(NativeConstants.SM_CXDOUBLECLK));
        _doubleClickHeight = Math.Max(1, NativeMethods.GetSystemMetrics(NativeConstants.SM_CYDOUBLECLK));
        if (settings.CloseTabOnDoubleClick && !SetEnabled(true))
            settings.CloseTabOnDoubleClick = false;
    }

    public bool SetEnabled(bool enabled)
    {
        lock (_sync)
        {
            if (_disposed)
                return false;

            if (enabled == _enabled)
                return _enabled;

            if (enabled)
            {
                if (!InstallHook())
                    return false;

                _enabled = true;
                _logger.Info("Explorer tab close-on-double-click enabled.");
                return true;
            }

            UninstallHook();
            ResetClickState();
            _enabled = false;
            _logger.Info("Explorer tab close-on-double-click disabled.");
            return false;
        }
    }

    private bool InstallHook()
    {
        _hookHandle = NativeMethods.SetWindowsHookEx(
            NativeConstants.WH_MOUSE_LL,
            _mouseProc,
            IntPtr.Zero,
            0);

        if (_hookHandle == IntPtr.Zero)
        {
            IntPtr module = NativeMethods.GetModuleHandle(null);
            if (module != IntPtr.Zero)
            {
                _hookHandle = NativeMethods.SetWindowsHookEx(
                    NativeConstants.WH_MOUSE_LL,
                    _mouseProc,
                    module,
                    0);
            }
        }

        if (_hookHandle != IntPtr.Zero)
            return true;

        int lastError = Marshal.GetLastWin32Error();
        _logger.Warn($"Failed to install Explorer mouse hook. Win32Error={lastError}");
        _hookHandle = IntPtr.Zero;
        return false;
    }

    private void UninstallHook()
    {
        if (_hookHandle == IntPtr.Zero)
            return;

        if (!NativeMethods.UnhookWindowsHookEx(_hookHandle))
        {
            int lastError = Marshal.GetLastWin32Error();
            _logger.Warn($"Failed to uninstall Explorer mouse hook. Win32Error={lastError}");
        }

        _hookHandle = IntPtr.Zero;
    }

    private IntPtr MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode < 0 || wParam != (IntPtr)NativeConstants.WM_LBUTTONDOWN || lParam == IntPtr.Zero)
            return NativeMethods.CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);

        try
        {
            if (HandleLeftButtonDown(lParam))
            {
                // Eat the double-click so Explorer doesn't execute its native action (e.g., maximize window)
                return new IntPtr(1);
            }
        }
        catch (Exception ex)
        {
            _logger.Debug($"Explorer mouse hook callback failed: {ex.Message}");
        }

        return NativeMethods.CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
    }

    private bool HandleLeftButtonDown(IntPtr lParam)
    {
        var info = Marshal.PtrToStructure<NativeStructs.MSLLHOOKSTRUCT>(lParam);
        if (!TryResolveExplorerTabByPoint(info.pt, out IntPtr topLevelHandle, out IntPtr tabHandle))
        {
            lock (_sync) { ResetClickState(); }
            return false;
        }

        lock (_sync)
        {
            if (_disposed || !_enabled)
                return false;

            if (!IsPointInTabTitleArea(info.pt, topLevelHandle))
            {
                ResetClickState();
                return false;
            }

            long now = Environment.TickCount64;
            bool isDoubleClick =
                _lastClickTopLevelHandle == topLevelHandle &&
                now - _lastClickTick >= 0 &&
                now - _lastClickTick <= _doubleClickTimeMs &&
                Math.Abs(info.pt.X - _lastClickPoint.X) <= _doubleClickWidth &&
                Math.Abs(info.pt.Y - _lastClickPoint.Y) <= _doubleClickHeight;

            _lastClickTick = now;
            _lastClickTopLevelHandle = topLevelHandle;
            _lastClickPoint = info.pt;

            if (!isDoubleClick)
                return false;

            ResetClickState();
            _logger.Info("Explorer tab double-click detected; sending close command.");
            CloseTab(topLevelHandle, tabHandle);
            return true;
        }
    }

    private static bool TryResolveExplorerTabByPoint(
        NativeStructs.POINT point,
        out IntPtr topLevelHandle,
        out IntPtr tabHandle)
    {
        topLevelHandle = IntPtr.Zero;
        tabHandle = IntPtr.Zero;

        IntPtr pointWindow = NativeMethods.WindowFromPoint(point);
        if (pointWindow == IntPtr.Zero)
            return false;

        topLevelHandle = NativeMethods.GetAncestor(pointWindow, NativeConstants.GA_ROOT);
        if (topLevelHandle == IntPtr.Zero || !WindowClassEquals(topLevelHandle, ExplorerWindowClass))
            return false;

        tabHandle = TryFindAncestorWindowByClass(pointWindow, ExplorerTabClass, stopAt: topLevelHandle);
        if (tabHandle == IntPtr.Zero || !NativeMethods.IsWindow(tabHandle))
        {
            tabHandle = NativeMethods.FindWindowEx(topLevelHandle, IntPtr.Zero, ExplorerTabClass, null);
        }

        return tabHandle != IntPtr.Zero && NativeMethods.IsWindow(tabHandle);
    }

    /// <summary>
    /// Returns true only when the click is on an actual tab title, not on empty caption space.
    ///
    /// In Windows 11 Explorer the tab-bar area is split between two child window classes:
    ///
    ///   • Microsoft.UI.Content.DesktopChildSiteBridge — hosts the WinUI 3 tab items and
    ///     buttons.  A click here is always on a real UI element (tab title, close button,
    ///     new-tab button, etc.).
    ///
    ///   • TITLE_BAR_SCAFFOLDING_WINDOW_CLASS — covers the empty/draggable portion of the
    ///     title bar to the right of the last tab.  A double-click here should trigger the
    ///     native maximize/restore, not close a tab.
    ///
    /// We accept clicks that land on the WinUI bridge window and are geometrically within
    /// the vertical tab-strip band (above the ShellTabWindowClass content area).
    /// </summary>
    private bool IsPointInTabTitleArea(NativeStructs.POINT point, IntPtr topLevelHandle)
    {
        IntPtr pointWindow = NativeMethods.WindowFromPoint(point);
        string wfpClass = GetWindowClass(pointWindow);


        // If the cursor is over the title-bar scaffolding (empty draggable space), reject
        // immediately so the OS can handle native double-click (maximize / restore).
        if (string.Equals(wfpClass, TitleBarScaffoldingClass, StringComparison.OrdinalIgnoreCase))
            return false;

        // Accept only clicks on the WinUI bridge that hosts the tab items.
        // Reject anything else (top-level frame, address bar, toolbar, etc.).
        if (!string.Equals(wfpClass, WinUiBridgeClass, StringComparison.OrdinalIgnoreCase))
            return false;

        // Also verify the point is within the vertical tab-strip band (above ShellTabWindowClass).
        if (!NativeMethods.GetWindowRect(topLevelHandle, out NativeStructs.RECT explorerRect))
            return false;

        IntPtr tabHandle = NativeMethods.FindWindowEx(topLevelHandle, IntPtr.Zero, ExplorerTabClass, null);
        if (tabHandle == IntPtr.Zero || !NativeMethods.GetWindowRect(tabHandle, out NativeStructs.RECT tabRect))
            return false;

        int captionHeight = Math.Max(0, NativeMethods.GetSystemMetrics(NativeConstants.SM_CYCAPTION));
        int frameHeight   = Math.Max(0, NativeMethods.GetSystemMetrics(NativeConstants.SM_CYFRAME));
        int maxTabStripHeight = Math.Clamp(captionHeight + frameHeight + 8, 30, 56);

        int stripTop    = explorerRect.Top;
        int stripBottom = Math.Min(tabRect.Top, stripTop + maxTabStripHeight);
        if (stripBottom <= stripTop)
            stripBottom = stripTop + maxTabStripHeight;

        bool inStrip = point.Y >= stripTop && point.Y < stripBottom;
        return inStrip;
    }

    private static string GetWindowClass(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero) return string.Empty;
        var sb = new StringBuilder(128);
        NativeMethods.GetClassName(hwnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static IntPtr TryFindAncestorWindowByClass(IntPtr hwnd, string expectedClass, IntPtr stopAt)
    {
        IntPtr current = hwnd;
        for (int depth = 0; depth < 24 && current != IntPtr.Zero; depth++)
        {
            if (WindowClassEquals(current, expectedClass))
                return current;

            if (current == stopAt)
                break;

            current = NativeMethods.GetParent(current);
        }

        return IntPtr.Zero;
    }

    private static bool WindowClassEquals(IntPtr hwnd, string expectedClass)
    {
        if (hwnd == IntPtr.Zero)
            return false;

        var className = new StringBuilder(64);
        NativeMethods.GetClassName(hwnd, className, className.Capacity);
        return string.Equals(className.ToString(), expectedClass, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Closes the currently active tab by sending WM_COMMAND with EXPLORER_CMD_CLOSE_TAB.
    /// Prefer the resolved tab handle and fall back to top-level Explorer window when needed.
    /// </summary>
    private static void CloseTab(IntPtr topLevelHandle, IntPtr tabHandle)
    {
        IntPtr commandTarget = IntPtr.Zero;
        if (tabHandle != IntPtr.Zero && NativeMethods.IsWindow(tabHandle))
            commandTarget = tabHandle;
        else if (topLevelHandle != IntPtr.Zero && NativeMethods.IsWindow(topLevelHandle))
            commandTarget = topLevelHandle;

        if (commandTarget == IntPtr.Zero)
            return;

        NativeMethods.PostMessage(
            commandTarget,
            (uint)NativeConstants.WM_COMMAND,
            new IntPtr(NativeConstants.EXPLORER_CMD_CLOSE_TAB),
            new IntPtr(1));
    }

    private void ResetClickState()
    {
        _lastClickTick = 0;
        _lastClickTopLevelHandle = IntPtr.Zero;
        _lastClickPoint = default;
    }

    public void Dispose()
    {
        lock (_sync)
        {
            if (_disposed)
                return;

            _disposed = true;
            UninstallHook();
            ResetClickState();
            _enabled = false;
        }
    }
}
