using WinTab.Hooks;

namespace WinTab.Models;

internal sealed class WindowSearchItem
{
    public nint Handle { get; set; }
    public string Name { get; set; } = string.Empty;
    public string DisplayLocation { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public string[]? SelectedItems { get; set; }
    public WindowHostType HostType { get; set; } = WindowHostType.Unknown;
}

