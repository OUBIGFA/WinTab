using System.Runtime.InteropServices;

namespace WinTab.ShellBridge.Interop;

internal static class User32Native
{
    [DllImport("user32.dll")]
    internal static extern nint GetForegroundWindow();
}
