namespace WinTab.Core.Models;

/// <summary>
/// Serialized descriptor for one tab in a saved group session.
/// Used to find matching windows on the next startup.
/// </summary>
public sealed class GroupWindowTabState
{
    public int Order { get; set; }
    public string ProcessName { get; set; } = string.Empty;
    public string WindowTitle { get; set; } = string.Empty;
    public string ClassName { get; set; } = string.Empty;
    public string? ProcessPath { get; set; }
}
