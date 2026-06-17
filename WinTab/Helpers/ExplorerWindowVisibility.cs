using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using WinTab.WinAPI;

namespace WinTab.Helpers;

public static class ExplorerWindowVisibility
{
    private const int SM_XVIRTUALSCREEN = 76;
    private const int SM_YVIRTUALSCREEN = 77;
    private const int SM_CXVIRTUALSCREEN = 78;
    private const int SM_CYVIRTUALSCREEN = 79;
    private const int OffscreenRestoreMargin = 120;

    private static readonly ConcurrentDictionary<nint, RECT?> HiddenWindows = new();

    public static IEnumerable<nint> HiddenWindowHandles => HiddenWindows.Keys;

    public static bool Contains(nint hWnd) => HiddenWindows.ContainsKey(hWnd);

    public static bool Forget(nint hWnd) => HiddenWindows.TryRemove(hWnd, out _);

    public static void Hide(nint hWnd, bool keepTheme = false)
    {
        var originalPos = HiddenWindows.GetOrAdd(hWnd, static (hWnd, keepTheme) =>
        {
            if (!keepTheme)
                return null;

            return WinApi.GetWindowRect(hWnd, out var originalPos) ? originalPos : null;
        }, keepTheme);

        if (!keepTheme)
        {
            UpdateWindowLayered(hWnd, remove: false);
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 0, WinApi.LWA_ALPHA);
            return;
        }

        if (originalPos == null && WinApi.GetWindowRect(hWnd, out var currentPos))
        {
            originalPos = currentPos;
            HiddenWindows[hWnd] = currentPos;
        }

        const uint flags = WinApi.SWP_HIDEWINDOW | WinApi.SWP_NOSIZE | WinApi.SWP_NOZORDER | WinApi.SWP_NOACTIVATE | WinApi.SWP_FRAMECHANGED;
        WinApi.SetWindowPos(hWnd, 0, -32_000, -32_000, 0, 0, flags);
    }

    public static bool Show(nint hWnd, bool removeCache)
    {
        var restored = Restore(hWnd, removeCache, removeLayeredStyle: false);
        if (restored)
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 255, WinApi.LWA_ALPHA);

        return restored;
    }

    public static int RestoreAll(bool removeLayeredStyle = true)
    {
        var restored = 0;
        var candidates = ExplorerWindowDiscovery.GetAllExplorerWindows()
            .Concat(HiddenWindows.Keys)
            .Distinct()
            .ToArray();

        foreach (var hWnd in candidates)
        {
            if (Restore(hWnd, removeCache: true, removeLayeredStyle))
                restored++;
        }

        return restored;
    }

    public static bool Restore(nint hWnd, bool removeCache = true, bool removeLayeredStyle = true)
    {
        if (hWnd == 0 || !ExplorerWindowDiscovery.IsFileExplorerWindow(hWnd))
        {
            if (removeCache)
                Forget(hWnd);

            return false;
        }

        var hasCache = HiddenWindows.TryGetValue(hWnd, out var originalPos);
        var exStyle = WinApi.GetWindowLong(hWnd, WinApi.GWL_EXSTYLE);
        var isLayered = (exStyle & WinApi.WS_EX_LAYERED) != 0;
        var alpha = (byte)255;
        var alphaFlags = 0u;
        var isTransparent = isLayered &&
                            WinApi.GetLayeredWindowAttributes(hWnd, out _, out alpha, out alphaFlags) &&
                            (alphaFlags & (uint)WinApi.LWA_ALPHA) != 0 &&
                            alpha == 0;
        var isVisible = WinApi.IsWindowVisible(hWnd);
        var hasRect = WinApi.GetWindowRect(hWnd, out var rect);
        var isOffscreen = hasRect && IsExplorerWindowOffscreen(rect);

        if (!hasCache && !isTransparent && isVisible && !isOffscreen)
            return false;

        if (removeCache)
            Forget(hWnd);

        const uint showFlags = WinApi.SWP_SHOWWINDOW | WinApi.SWP_NOSIZE | WinApi.SWP_NOZORDER | WinApi.SWP_NOACTIVATE | WinApi.SWP_FRAMECHANGED;
        if (originalPos != null)
        {
            WinApi.SetWindowPos(hWnd, 0, originalPos.Value.Left, originalPos.Value.Top, 0, 0, showFlags);
        }
        else if (isOffscreen)
        {
            var x = WinApi.GetSystemMetrics(SM_XVIRTUALSCREEN) + OffscreenRestoreMargin;
            var y = WinApi.GetSystemMetrics(SM_YVIRTUALSCREEN) + OffscreenRestoreMargin;
            WinApi.SetWindowPos(hWnd, 0, x, y, 0, 0, showFlags);
        }
        else
        {
            WinApi.ShowWindow(hWnd, WinApi.SW_SHOWNOACTIVATE);
            if (hasRect)
                WinApi.SetWindowPos(hWnd, 0, rect.Left, rect.Top, 0, 0, showFlags);
        }

        if (isLayered)
        {
            WinApi.SetLayeredWindowAttributes(hWnd, 0, 255, WinApi.LWA_ALPHA);
            if (removeLayeredStyle)
                UpdateWindowLayered(hWnd, remove: true);
        }

        return true;
    }

    public static void UpdateLayeredStyle(nint hWnd, bool remove)
    {
        UpdateWindowLayered(hWnd, remove);
    }

    private static void UpdateWindowLayered(nint hWnd, bool remove)
    {
        var exStyle = WinApi.GetWindowLong(hWnd, WinApi.GWL_EXSTYLE);
        var isLayered = (exStyle & WinApi.WS_EX_LAYERED) != 0;

        if (remove && isLayered)
            WinApi.SetWindowLong(hWnd, WinApi.GWL_EXSTYLE, exStyle & ~WinApi.WS_EX_LAYERED);

        if (!remove && !isLayered)
            WinApi.SetWindowLong(hWnd, WinApi.GWL_EXSTYLE, exStyle | WinApi.WS_EX_LAYERED);
    }

    private static bool IsExplorerWindowOffscreen(RECT rect)
    {
        var width = rect.Right - rect.Left;
        var height = rect.Bottom - rect.Top;
        if (width < 80 || height < 80)
            return false;

        var virtualLeft = WinApi.GetSystemMetrics(SM_XVIRTUALSCREEN);
        var virtualTop = WinApi.GetSystemMetrics(SM_YVIRTUALSCREEN);
        var virtualRight = virtualLeft + WinApi.GetSystemMetrics(SM_CXVIRTUALSCREEN);
        var virtualBottom = virtualTop + WinApi.GetSystemMetrics(SM_CYVIRTUALSCREEN);

        return rect.Right < virtualLeft - OffscreenRestoreMargin ||
               rect.Bottom < virtualTop - OffscreenRestoreMargin ||
               rect.Left > virtualRight + OffscreenRestoreMargin ||
               rect.Top > virtualBottom + OffscreenRestoreMargin;
    }
}
