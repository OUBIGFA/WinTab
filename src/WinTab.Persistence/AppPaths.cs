namespace WinTab.Persistence;

/// <summary>
/// Provides well-known file and directory paths used by the application.
/// If a <c>portable.txt</c> file exists next to the executable, portable mode is enabled
/// and all data is stored in a <c>data</c> subdirectory next to the executable.
/// Otherwise, data is stored in <c>%AppData%/WinTab</c>.
/// </summary>
public static class AppPaths
{
    private static readonly string PortableMarkerFileName = "portable.txt";
    private static readonly string DataFolderName = "data";

    private static readonly bool IsPortableMode;
    private static readonly string AppBaseDirectory = 
        Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory;

    static AppPaths()
    {
        string portableMarkerPath = Path.Combine(AppBaseDirectory, PortableMarkerFileName);
        IsPortableMode = File.Exists(portableMarkerPath);
    }

    /// <summary>
    /// Root directory for all WinTab application data.
    /// In portable mode: executable directory + "data" subfolder.
    /// In installed mode: <c>%AppData%\WinTab</c>.
    /// </summary>
    public static string BaseDirectory { get; } =
        IsPortableMode
            ? Path.Combine(AppBaseDirectory, DataFolderName)
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WinTab");

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

    /// <summary>
    /// Indicates whether the application is running in portable mode.
    /// </summary>
    public static bool IsPortable => IsPortableMode;
}
