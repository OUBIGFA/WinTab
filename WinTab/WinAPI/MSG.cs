// ReSharper disable InconsistentNaming

using System.Runtime.InteropServices;

namespace WinTab.WinAPI;

[StructLayout(LayoutKind.Sequential)]
public struct MSG
{
    public nint hwnd;
    public uint message;
    public nint wParam;
    public nint lParam;
    public uint time;
    public int pt_x;
    public int pt_y;
}
