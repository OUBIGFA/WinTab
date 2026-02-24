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
    private const int RoleSystemPageTab = 0x25;
    private const int RoleSystemPageTabList = 0x3C;

    private readonly object _sync = new();
    private readonly Logger _logger;
    private readonly NativeMethods.LowLevelMouseProc _mouseProc;
    private readonly int _doubleClickTimeMs;
    private readonly int _doubleClickWidth;
    private readonly int _doubleClickHeight;
    private readonly int _tabHeaderFallbackHeight;

    private IntPtr _hookHandle;
    private bool _enabled;
    private bool _disposed;

    private long _lastClickTick;
    private IntPtr _lastClickTopLevelHandle;
    private IntPtr _lastClickTabHandle;
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
        _tabHeaderFallbackHeight = Math.Max(
            48,
            Math.Max(0, NativeMethods.GetSystemMetrics(NativeConstants.SM_CYCAPTION)) +
            Math.Max(0, NativeMethods.GetSystemMetrics(NativeConstants.SM_CYFRAME)) +
            40);

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
            HandleLeftButtonDown(lParam);
        }
        catch (Exception ex)
        {
            _logger.Debug($"Explorer mouse hook callback failed: {ex.Message}");
        }

        return NativeMethods.CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
    }

    private void HandleLeftButtonDown(IntPtr lParam)
    {
        var info = Marshal.PtrToStructure<NativeStructs.MSLLHOOKSTRUCT>(lParam);
        if (!TryResolveExplorerTabByPoint(info.pt, out IntPtr topLevelHandle, out IntPtr tabHandle))
        {
            lock (_sync)
            {
                ResetClickState();
            }

            return;
        }

        lock (_sync)
        {
            if (_disposed || !_enabled)
                return;

            if (!IsPointInTabTitleArea(info.pt, topLevelHandle))
            {
                ResetClickState();
                return;
            }

            long now = Environment.TickCount64;
            bool isDoubleClick =
                _lastClickTopLevelHandle == topLevelHandle &&
                _lastClickTabHandle == tabHandle &&
                now - _lastClickTick >= 0 &&
                now - _lastClickTick <= _doubleClickTimeMs &&
                Math.Abs(info.pt.X - _lastClickPoint.X) <= _doubleClickWidth &&
                Math.Abs(info.pt.Y - _lastClickPoint.Y) <= _doubleClickHeight;

            _lastClickTick = now;
            _lastClickTopLevelHandle = topLevelHandle;
            _lastClickTabHandle = tabHandle;
            _lastClickPoint = info.pt;

            if (!isDoubleClick)
                return;

            ResetClickState();
            CloseTab(tabHandle);
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

        tabHandle = NativeMethods.FindWindowEx(topLevelHandle, IntPtr.Zero, ExplorerTabClass, null);
        return tabHandle != IntPtr.Zero && NativeMethods.IsWindow(tabHandle);
    }

    private bool IsPointInTabTitleArea(NativeStructs.POINT point, IntPtr topLevelHandle)
    {
        if (IsPointOnAccessibleTab(point))
            return true;

        if (!NativeMethods.GetWindowRect(topLevelHandle, out NativeStructs.RECT topLevelRect))
            return false;

        int maxHeaderY = topLevelRect.Top + _tabHeaderFallbackHeight;
        return point.Y >= topLevelRect.Top && point.Y <= maxHeaderY;
    }

    private static bool IsPointOnAccessibleTab(NativeStructs.POINT point)
    {
        if (NativeMethods.AccessibleObjectFromPoint(point, out Accessibility.IAccessible accObj, out object childId) != 0)
            return false;

        if (accObj is null)
            return false;

        try
        {
            int roleFromChild = RoleToInt(accObj.get_accRole(childId));
            if (roleFromChild is RoleSystemPageTab or RoleSystemPageTabList)
                return true;

            int roleFromSelf = RoleToInt(accObj.get_accRole(0));
            return roleFromSelf is RoleSystemPageTab or RoleSystemPageTabList;
        }
        catch
        {
            return false;
        }
        finally
        {
            try
            {
                Marshal.FinalReleaseComObject(accObj);
            }
            catch
            {
                // ignore release failures
            }
        }
    }

    private static int RoleToInt(object roleValue)
    {
        return roleValue switch
        {
            int value => value,
            short value => value,
            byte value => value,
            _ => Convert.ToInt32(roleValue)
        };
    }

    private static bool WindowClassEquals(IntPtr hwnd, string expectedClass)
    {
        if (hwnd == IntPtr.Zero)
            return false;

        var className = new StringBuilder(64);
        NativeMethods.GetClassName(hwnd, className, className.Capacity);
        return string.Equals(className.ToString(), expectedClass, StringComparison.OrdinalIgnoreCase);
    }

    private static void CloseTab(IntPtr tabHandle)
    {
        if (tabHandle == IntPtr.Zero || !NativeMethods.IsWindow(tabHandle))
            return;

        NativeMethods.PostMessage(
            tabHandle,
            (uint)NativeConstants.WM_COMMAND,
            new IntPtr(NativeConstants.EXPLORER_CMD_CLOSE_TAB),
            new IntPtr(1));
    }

    private void ResetClickState()
    {
        _lastClickTick = 0;
        _lastClickTopLevelHandle = IntPtr.Zero;
        _lastClickTabHandle = IntPtr.Zero;
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
