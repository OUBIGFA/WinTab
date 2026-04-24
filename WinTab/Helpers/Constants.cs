namespace WinTab.Helpers;

internal static class Constants
{
    internal const string AppName = "WinTab";
    internal const string MutexId = $"__{AppName}Hook__Mutex";
    internal const string BackgroundLaunchArg = "--background";
    internal const string NotifyIconText = "WinTab keeps File Explorer folders in tabs.";
    internal const string SettingsFileName = "settings.json";
    internal const string UpdateUrl = "https://api.github.com/repos/OUBIGFA/WinTab/releases/latest";
}
