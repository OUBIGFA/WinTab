using System;
using System.IO;
using WinTab.Helpers;

namespace WinTab.Hooks;

internal static class ExplorerDebugLog
{
    private static readonly string? LogPath = Environment.GetEnvironmentVariable("WINTAB_DEBUG_LOG");
    private static readonly BufferedDiagnosticLog? Writer = string.IsNullOrWhiteSpace(LogPath)
        ? null
        : new BufferedDiagnosticLog(() => new FileStream(LogPath, FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite));

    public static void Write(string message)
    {
        Writer?.TryWrite(message);
    }

    public static bool Complete(TimeSpan timeout) => Writer?.CompleteAsync().Wait(timeout) ?? true;
}
