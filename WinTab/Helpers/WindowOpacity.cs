using WinTab.WinAPI;

namespace WinTab.Helpers;

/// <summary>The native opacity boundary, separate from ownership records and ordered taskbar recovery.</summary>
internal interface IWindowOpacity
{
    int ReadStyle(nint handle);
    void WriteStyle(nint handle, int style);
    bool TryRead(nint handle, out uint colorKey, out byte alpha, out uint flags);
    bool TryWrite(nint handle, uint colorKey, byte alpha, uint flags);
}

internal sealed class NativeWindowOpacity : IWindowOpacity
{
    internal static readonly NativeWindowOpacity Instance = new();
    private NativeWindowOpacity() { }
    public int ReadStyle(nint handle) => WinApi.GetWindowLong(handle, WinApi.GWL_EXSTYLE);
    public void WriteStyle(nint handle, int style) => WinApi.SetWindowLong(handle, WinApi.GWL_EXSTYLE, style);
    public bool TryRead(nint handle, out uint colorKey, out byte alpha, out uint flags) =>
        WinApi.GetLayeredWindowAttributes(handle, out colorKey, out alpha, out flags);
    public bool TryWrite(nint handle, uint colorKey, byte alpha, uint flags) =>
        WinApi.SetLayeredWindowAttributes(handle, colorKey, alpha, flags);
}
