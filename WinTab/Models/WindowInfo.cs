using System;
using SHDocVw;
using WinTab.Helpers;

namespace WinTab.Models;

public class WindowInfo
{
    internal WindowIdentity Identity { get; init; }
    internal WindowIdentity TabIdentity { get; set; }
    internal int Generation { get; init; }
    internal bool Closed { get; set; }
    internal string[]? SelectedItems { get; set; }
    public bool EventsHooked { get; set; }
    public nint HookedTopLevelHWnd { get; set; }
    public string? Location { get; set; }
    public DWebBrowserEvents2_OnQuitEventHandler? OnQuitHandler { get; set; }
    public DWebBrowserEvents2_NavigateComplete2EventHandler? OnNavigateHandler { get; set; }

    internal void RefreshSelection(Func<string[]?> readSelection, Func<bool> isCurrent)
    {
        if (!isCurrent())
            return;
        var location = Location;
        var selection = readSelection();
        if (selection != null && isCurrent() && Location == location)
            SelectedItems = selection;
    }
}
