using WinTab.Core.Models;

namespace WinTab.Core.Events;

public sealed class TabSwitchedEventArgs : EventArgs
{
    public required TabGroup Group { get; init; }
    public required TabItem PreviousTab { get; init; }
    public required TabItem NewTab { get; init; }
}
