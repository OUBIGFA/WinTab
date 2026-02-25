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
    private const int RoleSystemToolBar = 0x16;
    private const int RoleSystemPushButton = 0x2B;
    private const int RoleSystemSplitButton = 0x3E;
    private const int RoleSystemButtonDropDown = 0x38;
    private const int RoleSystemButtonMenu = 0x39;
    private const int RoleSystemButtonDropDownGrid = 0x3A;

    private enum AccessiblePointKind
    {
        Unknown,
        Tab,
        NavigationControl,
        Other
    }

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
            lock (_sync)
            {
                ResetClickState();
            }

            return false;
        }

        lock (_sync)
        {
            if (_disposed || !_enabled)
                return false;

            if (!IsPointInTabTitleArea(info.pt, topLevelHandle, tabHandle))
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
    /// Geometric check: the click point is in the tab title strip if its Y coordinate
    /// is above the top edge of the ShellTabWindowClass (the content area) but still
    /// within the CabinetWClass (the Explorer window). This is far more reliable than
    /// the Accessibility API for the Windows 11 Explorer tab bar, which lives in the
    /// custom title bar area and does not consistently report PageTab roles.
    /// </summary>
    private static bool IsPointInTabTitleArea(NativeStructs.POINT point, IntPtr topLevelHandle, IntPtr tabHandle)
    {
        AccessiblePointKind pointKind = GetAccessiblePointKind(point);
        if (pointKind == AccessiblePointKind.Tab)
            return true;

        if (pointKind == AccessiblePointKind.NavigationControl)
            return false;

        if (!NativeMethods.GetWindowRect(topLevelHandle, out NativeStructs.RECT explorerRect))
            return false;

        // Click must be within the Explorer window horizontally.
        if (point.X < explorerRect.Left || point.X > explorerRect.Right)
            return false;

        if (!NativeMethods.GetWindowRect(tabHandle, out NativeStructs.RECT tabRect))
            return false;

        int captionHeight = Math.Max(0, NativeMethods.GetSystemMetrics(NativeConstants.SM_CYCAPTION));
        int frameHeight = Math.Max(0, NativeMethods.GetSystemMetrics(NativeConstants.SM_CYFRAME));
        int maxTabStripHeight = Math.Clamp(captionHeight + frameHeight + 8, 30, 56);

        int headerTop = explorerRect.Top;
        int headerBottomExclusive = Math.Min(tabRect.Top, headerTop + maxTabStripHeight);

        if (headerBottomExclusive <= headerTop)
            headerBottomExclusive = headerTop + maxTabStripHeight;

        // The tab title strip sits above the ShellTabWindowClass content area.
        return point.Y >= headerTop && point.Y < headerBottomExclusive;
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

    private static AccessiblePointKind GetAccessiblePointKind(NativeStructs.POINT point)
    {
        if (NativeMethods.AccessibleObjectFromPoint(point, out Accessibility.IAccessible accObj, out object childId) != 0)
            return AccessiblePointKind.Unknown;

        if (accObj is null)
            return AccessiblePointKind.Unknown;

        try
        {
            int? roleFromChild = TryGetAccessibleRole(accObj, childId);
            int? roleFromSelf = TryGetAccessibleRole(accObj, 0);

            bool isTabLike =
                IsTabRole(roleFromChild) ||
                IsTabRole(roleFromSelf) ||
                (HasTabLikeAccessibleAncestor(accObj) &&
                 (IsTabButtonRole(roleFromChild) || IsTabButtonRole(roleFromSelf)));

            if (isTabLike)
                return AccessiblePointKind.Tab;

            bool isNavigationLike = IsNavigationLikeRole(roleFromChild) || IsNavigationLikeRole(roleFromSelf);
            return isNavigationLike ? AccessiblePointKind.NavigationControl : AccessiblePointKind.Other;
        }
        catch
        {
            return AccessiblePointKind.Unknown;
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

    private static int? TryGetAccessibleRole(Accessibility.IAccessible accessible, object childId)
    {
        try
        {
            object roleValue = accessible.get_accRole(childId);
            return roleValue is null ? null : RoleToInt(roleValue);
        }
        catch
        {
            return null;
        }
    }

    private static bool HasTabLikeAccessibleAncestor(Accessibility.IAccessible accessible)
    {
        object? parentObj;
        try
        {
            parentObj = accessible.accParent;
        }
        catch
        {
            return false;
        }

        for (int depth = 0; depth < 8 && parentObj is not null; depth++)
        {
            if (parentObj is not Accessibility.IAccessible parentAccessible)
                return false;

            object? nextParent = null;

            try
            {
                int? parentRole = TryGetAccessibleRole(parentAccessible, 0);
                if (IsTabRole(parentRole))
                    return true;

                nextParent = parentAccessible.accParent;
            }
            catch
            {
                return false;
            }
            finally
            {
                try
                {
                    Marshal.FinalReleaseComObject(parentAccessible);
                }
                catch
                {
                }
            }

            parentObj = nextParent;
        }

        return false;
    }

    private static bool IsTabRole(int? role)
    {
        return role is RoleSystemPageTab or RoleSystemPageTabList;
    }

    private static bool IsTabButtonRole(int? role)
    {
        return role is
            RoleSystemPushButton or
            RoleSystemSplitButton or
            RoleSystemButtonDropDown or
            RoleSystemButtonMenu or
            RoleSystemButtonDropDownGrid;
    }

    private static bool IsNavigationLikeRole(int? role)
    {
        return role is
            RoleSystemToolBar or
            RoleSystemPushButton or
            RoleSystemSplitButton or
            RoleSystemButtonDropDown or
            RoleSystemButtonMenu or
            RoleSystemButtonDropDownGrid;
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
