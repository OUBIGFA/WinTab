using SHDocVw;
using System.Diagnostics;

namespace WinTab.Models;

public class WindowInfo
{
    public long CreatedAt { get; } = Stopwatch.GetTimestamp();
    public bool CanAutoMerge { get; set; }
    public int AutoMergeAttempts { get; set; }
    public bool EventsHooked { get; set; }
    public string? Location { get; set; }
    public string? Name { get; set; }
    public DWebBrowserEvents2_OnQuitEventHandler? OnQuitHandler { get; set; }
    public DWebBrowserEvents2_NavigateComplete2EventHandler? OnNavigateHandler { get; set; }
}
