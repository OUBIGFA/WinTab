namespace WinTab.Core;

public sealed record WindowInfo(
    IntPtr Handle,
    string Title,
    string ProcessName,
    int ProcessId,
    string ClassName,
    bool IsVisible)
{
    public string HandleHex => $"0x{Handle.ToInt64():X}";
}
