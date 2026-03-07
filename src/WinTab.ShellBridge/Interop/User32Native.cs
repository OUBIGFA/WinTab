using System.Text;
using System.Runtime.InteropServices;

namespace WinTab.ShellBridge.Interop;

internal static class User32Native
{
    [DllImport("user32.dll")]
    internal static extern nint GetForegroundWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern int GetClassName(nint hWnd, StringBuilder className, int maxCount);
}
