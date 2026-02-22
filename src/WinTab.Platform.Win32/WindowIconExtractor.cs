using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace WinTab.Platform.Win32;

/// <summary>
/// Extracts icons from windows and converts them to PNG byte arrays.
/// Uses multiple strategies: WM_GETICON, GetClassLongPtr, and SHGetFileInfo.
/// </summary>
public static class WindowIconExtractor
{
    /// <summary>
    /// Extract the icon for a given window handle as PNG bytes.
    /// </summary>
    /// <param name="hWnd">The window handle.</param>
    /// <returns>PNG byte array, or null if no icon could be extracted.</returns>
    public static byte[]? ExtractIcon(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero || !NativeMethods.IsWindow(hWnd))
            return null;

        IntPtr hIcon = TryGetIconHandle(hWnd);
        if (hIcon == IntPtr.Zero)
            return null;

        return ConvertIconToPng(hIcon);
    }

    // ─── Icon Handle Retrieval ──────────────────────────────────────────────

    /// <summary>
    /// Attempts to retrieve an icon handle using multiple strategies in order of preference.
    /// </summary>
    private static IntPtr TryGetIconHandle(IntPtr hWnd)
    {
        IntPtr hIcon;

        // Strategy 1: WM_GETICON with ICON_SMALL2 (high-DPI small icon).
        hIcon = NativeMethods.SendMessage(
            hWnd, (uint)NativeConstants.WM_GETICON,
            (IntPtr)NativeConstants.ICON_SMALL2, IntPtr.Zero);
        if (hIcon != IntPtr.Zero)
            return hIcon;

        // Strategy 2: WM_GETICON with ICON_SMALL.
        hIcon = NativeMethods.SendMessage(
            hWnd, (uint)NativeConstants.WM_GETICON,
            (IntPtr)NativeConstants.ICON_SMALL, IntPtr.Zero);
        if (hIcon != IntPtr.Zero)
            return hIcon;

        // Strategy 3: WM_GETICON with ICON_BIG.
        hIcon = NativeMethods.SendMessage(
            hWnd, (uint)NativeConstants.WM_GETICON,
            (IntPtr)NativeConstants.ICON_BIG, IntPtr.Zero);
        if (hIcon != IntPtr.Zero)
            return hIcon;

        // Strategy 4: GetClassLongPtr with GCLP_HICONSM (small class icon).
        hIcon = NativeMethods.GetClassLongPtr(hWnd, NativeConstants.GCLP_HICONSM);
        if (hIcon != IntPtr.Zero)
            return hIcon;

        // Strategy 5: GetClassLongPtr with GCLP_HICON (large class icon).
        hIcon = NativeMethods.GetClassLongPtr(hWnd, NativeConstants.GCLP_HICON);
        if (hIcon != IntPtr.Zero)
            return hIcon;

        // Strategy 6: SHGetFileInfo using the process executable path.
        hIcon = TryGetIconFromProcess(hWnd);
        return hIcon;
    }

    /// <summary>
    /// Attempts to get an icon from the window's process executable via SHGetFileInfo.
    /// </summary>
    private static IntPtr TryGetIconFromProcess(IntPtr hWnd)
    {
        NativeMethods.GetWindowThreadProcessId(hWnd, out uint pid);
        if (pid == 0)
            return IntPtr.Zero;

        // Get the process executable path.
        IntPtr hProcess = NativeMethods.OpenProcess(
            NativeConstants.PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
        if (hProcess == IntPtr.Zero)
            return IntPtr.Zero;

        try
        {
            var buffer = new System.Text.StringBuilder(1024);
            uint size = (uint)buffer.Capacity;

            if (!NativeMethods.QueryFullProcessImageName(hProcess, 0, buffer, ref size))
                return IntPtr.Zero;

            string exePath = buffer.ToString();

            var shFileInfo = new NativeStructs.SHFILEINFO();
            IntPtr result = NativeMethods.SHGetFileInfo(
                exePath, 0, ref shFileInfo,
                (uint)Marshal.SizeOf<NativeStructs.SHFILEINFO>(),
                NativeConstants.SHGFI_ICON | NativeConstants.SHGFI_SMALLICON);

            if (result != IntPtr.Zero && shFileInfo.hIcon != IntPtr.Zero)
                return shFileInfo.hIcon;

            return IntPtr.Zero;
        }
        finally
        {
            NativeMethods.CloseHandle(hProcess);
        }
    }

    // ─── Icon-to-PNG Conversion ─────────────────────────────────────────────

    /// <summary>
    /// Converts an HICON to PNG byte array using System.Drawing.
    /// </summary>
    private static byte[]? ConvertIconToPng(IntPtr hIcon)
    {
        try
        {
            using var icon = Icon.FromHandle(hIcon);
            using var bitmap = icon.ToBitmap();
            using var ms = new MemoryStream();

            bitmap.Save(ms, ImageFormat.Png);
            return ms.ToArray();
        }
        catch (Exception)
        {
            // Icon handle might be invalid or conversion fails.
            return null;
        }
    }
}
