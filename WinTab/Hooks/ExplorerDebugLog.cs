using System;
using System.IO;

namespace WinTab.Hooks;

internal static class ExplorerDebugLog
{
    private static readonly string? LogPath = Environment.GetEnvironmentVariable("WINTAB_DEBUG_LOG");

    public static void Write(string message)
    {
        if (string.IsNullOrWhiteSpace(LogPath))
            return;

        try
        {
            File.AppendAllText(LogPath, $"{DateTime.Now:HH:mm:ss.fff} {message}{Environment.NewLine}");
        }
        catch
        {
            // 忽略调试日志写入失败，避免影响主流程。
        }
    }
}
