using WinTab.Models;

namespace WinTab.Hooks;

public interface IExplorerWatcherSettings
{
    bool HaveThemeIssue { get; }
    bool RestorePreviousWindows { get; }
    bool SaveClosedHistory { get; }
    int DefaultExplorerLaunchId { get; }
    WindowRecord[]? ClosedWindows { get; set; }
}

internal sealed class DefaultExplorerWatcherSettings : IExplorerWatcherSettings
{
    public bool HaveThemeIssue => false;
    public bool RestorePreviousWindows => true;
    public bool SaveClosedHistory => true;
    public int DefaultExplorerLaunchId => 1;
    public WindowRecord[]? ClosedWindows { get; set; }
}
