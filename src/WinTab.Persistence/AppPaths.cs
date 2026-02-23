namespace WinTab.Persistence;

/// <summary>
/// Provides well-known file and directory paths used by the application.
/// All paths are rooted under <c>%AppData%/WinTab</c>.
/// </summary>
public static class AppPaths
{
    /// <summary>
    /// Root directory for all WinTab application data.
    /// Typically <c>%AppData%\WinTab</c>.
    /// </summary>
    public static string BaseDirectory { get; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WinTab");

    /// <summary>
    /// Full path to the JSON settings file.
    /// </summary>
    public static string SettingsPath { get; } =
        Path.Combine(BaseDirectory, "settings.json");

    /// <summary>
    /// Directory containing log files.
    /// </summary>
    public static string LogsDirectory { get; } =
        Path.Combine(BaseDirectory, "logs");

    /// <summary>
    /// Full path to the main application log file.
    /// </summary>
    public static string LogPath { get; } =
        Path.Combine(LogsDirectory, "wintab.log");

    /// <summary>
    /// Full path to the crash log file.
    /// </summary>
    public static string CrashLogPath { get; } =
        Path.Combine(LogsDirectory, "crash.log");
}
