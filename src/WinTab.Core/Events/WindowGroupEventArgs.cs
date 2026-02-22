using WinTab.Core.Models;

namespace WinTab.Core.Events;

public sealed class WindowGroupEventArgs : EventArgs
{
    public required TabGroup Group { get; init; }
    public IntPtr AffectedWindow { get; init; }
}
