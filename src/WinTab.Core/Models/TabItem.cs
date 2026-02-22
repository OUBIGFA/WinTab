namespace WinTab.Core.Models;

public sealed class TabItem
{
    public required IntPtr Handle { get; init; }
    public string Title { get; set; } = string.Empty;
    public string ProcessName { get; set; } = string.Empty;
    public byte[]? IconData { get; set; }
    public string? AccentColor { get; set; }
    public bool IsPinned { get; set; }
    public int Order { get; set; }
}
