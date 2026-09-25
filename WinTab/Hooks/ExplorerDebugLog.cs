using System;
using System.IO;
using WinTab.Helpers;

namespace WinTab.Hooks;

internal static class ExplorerDebugLog
{
    private static readonly string? LogPath = Environment.GetEnvironmentVariable("WINTAB_DEBUG_LOG");
    private static readonly BufferedDiagnosticLog? Writer = string.IsNullOrWhiteSpace(LogPath)
        ? null
        : new BufferedDiagnosticLog(() => OpenLogFile(LogPath));

    public static void Write(string message)
    {
        Writer?.TryWrite(message);
    }

    public static bool Complete(TimeSpan timeout) => Writer?.CompleteAsync().Wait(timeout) ?? true;

    /// <summary>Opens the configured log file. A folder that does not exist yet is created rather than silently disabling the log.</summary>
    internal static Stream OpenLogFile(string path)
    {
        var folder = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(folder))
            Directory.CreateDirectory(folder);
        return new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);
    }
}
