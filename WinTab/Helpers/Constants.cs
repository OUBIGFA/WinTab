namespace WinTab.Helpers;

internal static class Constants
{
    internal const string AppName = "WinTab";
    internal const string MutexId = $"__{AppName}Hook__Mutex";
    internal const string SettingsFileName = "settings.json";
    internal const string HotKeyProfilesFileName = "HotKeyProfiles.json";
    internal const string JsonFileFilter = "JSON files (*.json)|*.json|All Files|*.*";
    internal const string UpdateUrl = "https://api.github.com/repos/w4po/WinTab/releases/latest";
    internal const string DefaultHotKeyProfiles = "[{\"Name\":\"Home\",\"HotKeys\":[91,69],\"Scope\":0,\"Action\":0,\"Path\":\"\",\"IsHandled\":true,\"IsEnabled\":true,\"Delay\":0},{\"Name\":\"Duplicate\",\"HotKeys\":[17,68],\"Scope\":1,\"Action\":1,\"Path\":null,\"IsHandled\":true,\"IsEnabled\":true,\"Delay\":0},{\"Name\":\"ReopenClosed\",\"HotKeys\":[16,17,84],\"Scope\":1,\"Action\":2,\"Path\":null,\"IsHandled\":true,\"IsEnabled\":true,\"Delay\":0}]";
    internal const int ShellReadyTimeoutMs = 2500;
    internal const int WindowHideDebounceMs = 250;
    internal const int OpenTabTimeoutMs = 2000;
    internal const int NavigateTimeoutMs = 5000;
    internal const int HealthCheckIntervalMs = 3000;
    internal const int ShellFailureThreshold = 3;
    internal const int ShellResetCooldownMs = 10000;
    internal const int ComRetryAttempts = 2;
    internal const int ComRetryDelayMs = 50;
    internal const int ComFailureThreshold = 5;
    internal const int ComFailureWindowMs = 5000;
}

