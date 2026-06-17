using WinTab.Hooks;
using WinTab.Models;

namespace WinTab.Managers;

internal sealed class ExplorerWatcherSettings : IExplorerWatcherSettings
{
    public static ExplorerWatcherSettings Instance { get; } = new();

    private ExplorerWatcherSettings()
    {
    }

    public bool HaveThemeIssue => SettingsManager.HaveThemeIssue;
    public bool RestorePreviousWindows => SettingsManager.RestorePreviousWindows;
    public bool SaveClosedHistory => SettingsManager.SaveClosedHistory;
    public int DefaultExplorerLaunchId => RegistryManager.GetDefaultExplorerLaunchId();

    public WindowRecord[]? ClosedWindows
    {
        get => SettingsManager.ClosedWindows;
        set => SettingsManager.ClosedWindows = value;
    }
}
