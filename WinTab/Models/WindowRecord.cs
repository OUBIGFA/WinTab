using System;

namespace WinTab.Models;

/// <summary>
/// In-memory note of a folder a user recently closed or asked to open, kept briefly so a
/// re-open of the same folder can restore its selection.
/// </summary>
public class WindowRecord(string location, nint handle = 0, string[]? selectedItems = null)
{
    public nint Handle { get; } = handle;
    public string Location { get; } = location;
    public string[]? SelectedItems { get; } = selectedItems;
    public long CreatedAt { get; } = Environment.TickCount;
}
