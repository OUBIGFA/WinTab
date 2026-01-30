namespace WinTab.App;

public sealed class GroupEntry
{
    public string Name { get; init; } = string.Empty;
    public bool IsOpen { get; init; }
    public int WindowCount { get; init; }
}
