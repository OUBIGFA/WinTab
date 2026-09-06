using SHDocVw;
using WinTab.Helpers;

namespace WinTab.Models;

public class WindowInfo
{
    internal WindowIdentity Identity { get; init; }
    internal int Generation { get; init; }
    internal bool Closed { get; set; }
    internal string[]? SelectedItems { get; set; }
    public bool EventsHooked { get; set; }
    public nint HookedTopLevelHWnd { get; set; }
    public string? Location { get; set; }
    public DWebBrowserEvents2_OnQuitEventHandler? OnQuitHandler { get; set; }
    public DWebBrowserEvents2_NavigateComplete2EventHandler? OnNavigateHandler { get; set; }
}
