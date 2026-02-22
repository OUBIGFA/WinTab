namespace WinTab.Platform.Win32;

/// <summary>
/// Static helpers for Desktop Window Manager (DWM) visual features.
/// Supports dark mode, Mica/Acrylic backdrops, and rounded corners on Windows 11+.
/// </summary>
public static class DwmHelper
{
    public static bool SetCloak(IntPtr hwnd, bool cloak)
    {
        int value = cloak ? 1 : 0;
        int hr = NativeMethods.DwmSetWindowAttribute(
            hwnd,
            NativeConstants.DWMWA_CLOAK,
            ref value,
            sizeof(int));
        return hr == 0;
    }

    /// <summary>
    /// Enable or disable the immersive dark mode title bar.
    /// Requires Windows 10 1809+ (build 17763) or later.
    /// </summary>
    /// <param name="hwnd">Target window handle.</param>
    /// <param name="dark">True to enable dark title bar; false for light.</param>
    /// <returns>True if the attribute was set successfully.</returns>
    public static bool SetDarkMode(IntPtr hwnd, bool dark)
    {
        int value = dark ? 1 : 0;
        int hr = NativeMethods.DwmSetWindowAttribute(
            hwnd,
            NativeConstants.DWMWA_USE_IMMERSIVE_DARK_MODE,
            ref value,
            sizeof(int));
        return hr == 0; // S_OK
    }

    /// <summary>
    /// Apply the Mica backdrop to a window (Windows 11 22H2+).
    /// Uses DWMWA_SYSTEMBACKDROP_TYPE = 38, value = DWMSBT_MAINWINDOW (2).
    /// </summary>
    /// <param name="hwnd">Target window handle.</param>
    /// <returns>True if the attribute was set successfully.</returns>
    public static bool SetMicaBackdrop(IntPtr hwnd)
    {
        int value = NativeConstants.DWMSBT_MAINWINDOW; // Mica
        int hr = NativeMethods.DwmSetWindowAttribute(
            hwnd,
            NativeConstants.DWMWA_SYSTEMBACKDROP_TYPE,
            ref value,
            sizeof(int));
        return hr == 0;
    }

    /// <summary>
    /// Apply the Acrylic backdrop to a window (Windows 11 22H2+).
    /// Uses DWMWA_SYSTEMBACKDROP_TYPE = 38, value = DWMSBT_TRANSIENTWINDOW (3).
    /// </summary>
    /// <param name="hwnd">Target window handle.</param>
    /// <returns>True if the attribute was set successfully.</returns>
    public static bool SetAcrylicBackdrop(IntPtr hwnd)
    {
        int value = NativeConstants.DWMSBT_TRANSIENTWINDOW; // Acrylic
        int hr = NativeMethods.DwmSetWindowAttribute(
            hwnd,
            NativeConstants.DWMWA_SYSTEMBACKDROP_TYPE,
            ref value,
            sizeof(int));
        return hr == 0;
    }

    /// <summary>
    /// Set or remove rounded corners on a window (Windows 11+).
    /// Uses DWMWA_WINDOW_CORNER_PREFERENCE = 33.
    /// </summary>
    /// <param name="hwnd">Target window handle.</param>
    /// <param name="rounded">True for rounded corners, false for square (do-not-round).</param>
    /// <returns>True if the attribute was set successfully.</returns>
    public static bool SetRoundedCorners(IntPtr hwnd, bool rounded)
    {
        int value = rounded ? NativeConstants.DWMWCP_ROUND : NativeConstants.DWMWCP_DONOTROUND;
        int hr = NativeMethods.DwmSetWindowAttribute(
            hwnd,
            NativeConstants.DWMWA_WINDOW_CORNER_PREFERENCE,
            ref value,
            sizeof(int));
        return hr == 0;
    }

    /// <summary>
    /// Determines whether the current OS is Windows 11 (build 22000) or later.
    /// </summary>
    /// <returns>True if running on Windows 11+.</returns>
    public static bool IsWindows11OrLater()
    {
        // Environment.OSVersion on .NET 5+ returns the true OS version.
        var version = Environment.OSVersion.Version;

        // Windows 11 starts at build 22000 (major 10, minor 0).
        return version.Major > 10 ||
               (version.Major == 10 && version.Build >= 22000);
    }

    /// <summary>
    /// Extends the DWM frame into the client area.
    /// Useful for enabling Mica/Acrylic effects on the entire window.
    /// Pass margins of -1 to extend to the full window.
    /// </summary>
    /// <param name="hwnd">Target window handle.</param>
    /// <param name="left">Left margin (-1 to extend fully).</param>
    /// <param name="right">Right margin.</param>
    /// <param name="top">Top margin.</param>
    /// <param name="bottom">Bottom margin.</param>
    /// <returns>True if the call succeeded.</returns>
    public static bool ExtendFrameIntoClientArea(IntPtr hwnd,
        int left = -1, int right = -1, int top = -1, int bottom = -1)
    {
        var margins = new NativeStructs.MARGINS
        {
            cxLeftWidth = left,
            cxRightWidth = right,
            cyTopHeight = top,
            cyBottomHeight = bottom
        };
        int hr = NativeMethods.DwmExtendFrameIntoClientArea(hwnd, ref margins);
        return hr == 0;
    }
}
