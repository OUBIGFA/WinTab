namespace WinTab.Core.Models;

public sealed record WindowInfo(
    IntPtr Handle,
    string Title,
    string ProcessName,
    int ProcessId,
    string ClassName,
    bool IsVisible,
    string? ProcessPath = null,
    byte[]? IconData = null)
{
    public string HandleHex => $"0x{Handle.ToInt64():X}";
}
