namespace WinTab.Persistence;

public static class AppPaths
{
    public static string BaseDirectory { get; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WinTab");

    public static string SettingsPath { get; } = Path.Combine(BaseDirectory, "settings.json");

    public static string LogsDirectory { get; } = Path.Combine(BaseDirectory, "logs");

    public static string LogPath { get; } = Path.Combine(LogsDirectory, "wintab.log");
}
