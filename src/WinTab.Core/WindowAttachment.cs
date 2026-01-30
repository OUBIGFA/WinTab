namespace WinTab.Core;

public sealed class WindowAttachment
{
    public required string GroupName { get; init; }

    public required IntPtr WindowHandle { get; init; }

    public int ProcessId { get; init; }

    public string? WindowTitle { get; init; }

    public string? ProcessName { get; init; }

    public DateTime AttachedAt { get; init; } = DateTime.UtcNow;
}
