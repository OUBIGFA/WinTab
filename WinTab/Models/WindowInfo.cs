using SHDocVw;

namespace WinTab.Models;

public class WindowInfo
{
    public bool EventsHooked { get; set; }
    public nint HookedTopLevelHWnd { get; set; }
    public string? Location { get; set; }
    public DWebBrowserEvents2_OnQuitEventHandler? OnQuitHandler { get; set; }
    public DWebBrowserEvents2_NavigateComplete2EventHandler? OnNavigateHandler { get; set; }
}
