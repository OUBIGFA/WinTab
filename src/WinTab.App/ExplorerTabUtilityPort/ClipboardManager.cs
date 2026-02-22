// This file is derived from ExplorerTabUtility (MIT License).
// Source: E:\_BIGFA Free\_code\ExplorerTabUtility

using System.Runtime.InteropServices;
using System.Text;

namespace WinTab.App.ExplorerTabUtilityPort;

internal static class ClipboardManager
{
    [DllImport("user32.dll")]
    private static extern bool OpenClipboard(nint hWndNewOwner);

    [DllImport("user32.dll")]
    private static extern bool CloseClipboard();

    [DllImport("user32.dll")]
    private static extern nint GetClipboardData(uint uFormat);

    [DllImport("user32.dll")]
    private static extern nint SetClipboardData(uint uFormat, nint hMem);

    [DllImport("user32.dll")]
    private static extern bool EmptyClipboard();

    [DllImport("kernel32.dll")]
    private static extern nint GlobalLock(nint hMem);

    [DllImport("kernel32.dll")]
    private static extern uint GlobalSize(nint hMem);

    [DllImport("kernel32.dll")]
    private static extern bool GlobalUnlock(nint hMem);

    [DllImport("kernel32.dll")]
    private static extern nint GlobalAlloc(uint uFlags, nuint dwBytes);

    [DllImport("kernel32.dll")]
    private static extern nint GlobalFree(nint hMem);

    private const uint GMEM_MOVEABLE = 0x0002;
    private const uint GMEM_ZEROINIT = 0x0040;
    private const uint GHND = GMEM_MOVEABLE | GMEM_ZEROINIT;
    private const uint CF_UNICODETEXT = 13;

    public static string GetClipboardText()
    {
        try
        {
            if (!OpenClipboard(default))
                return string.Empty;

            nint handle = GetClipboardData(CF_UNICODETEXT);
            if (handle == default)
                return string.Empty;

            nint lockedData = GlobalLock(handle);
            if (lockedData == default)
                return string.Empty;

            uint size = GlobalSize(handle);
            if (size == 0)
            {
                GlobalUnlock(handle);
                return string.Empty;
            }

            var buffer = new byte[size];
            Marshal.Copy(lockedData, buffer, 0, (int)size);
            GlobalUnlock(handle);

            return Encoding.Unicode.GetString(buffer).TrimEnd('\0');
        }
        catch
        {
            return string.Empty;
        }
        finally
        {
            CloseClipboard();
        }
    }

    public static void SetClipboardText(string text)
    {
        try
        {
            if (!OpenClipboard(default))
                return;

            EmptyClipboard();

            text = $"{text.TrimEnd('\0')}\0";
            byte[] buffer = Encoding.Unicode.GetBytes(text);
            uint size = (uint)buffer.Length;

            nint handle = GlobalAlloc(GHND, size);
            if (handle == default)
                return;

            nint pointer = GlobalLock(handle);
            if (pointer == default)
            {
                GlobalFree(handle);
                return;
            }

            Marshal.Copy(buffer, 0, pointer, (int)size);
            SetClipboardData(CF_UNICODETEXT, handle);
            GlobalUnlock(handle);
        }
        catch
        {
            // ignore
        }
        finally
        {
            CloseClipboard();
        }
    }
}
