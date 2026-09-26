using System;
using System.IO;
using WinTab.Helpers;

namespace WinTab.Hooks;

/// <summary>
/// WinTab's log. It is always on: every decision that can leave a folder, a tab or a click unhandled is written
/// here, so a problem that cannot be reproduced can still be read afterwards. WinTab writes one file per day
/// under <see cref="Folder"/> and keeps only the last two days.
/// </summary>
internal static class ExplorerDebugLog
{
    /// <summary>Tests and stress probes give the WinTab they run a file of its own, away from the user's log.</summary>
    internal const string FileOverrideVariable = "WINTAB_LOG_FILE";

    public static string Folder { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WinTab", "logs");

    private static readonly BufferedDiagnosticLog Writer = Create(Environment.GetEnvironmentVariable(FileOverrideVariable));

    private static BufferedDiagnosticLog Create(string? file) => string.IsNullOrWhiteSpace(file)
        ? new BufferedDiagnosticLog(new DailyLogSink(Folder))
        : new BufferedDiagnosticLog(() => OpenLogFile(file));

    public static void Write(string message) => Writer.TryWrite(message);

    public static bool Complete(TimeSpan timeout) => Writer.CompleteAsync().Wait(timeout);

    /// <summary>Opens the configured log file. A folder that does not exist yet is created rather than silently disabling the log.</summary>
    internal static Stream OpenLogFile(string path)
    {
        var folder = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(folder))
            Directory.CreateDirectory(folder);
        return new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);
    }
}
