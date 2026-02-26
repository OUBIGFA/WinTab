using System;
using System.IO;
using System.Reflection;

namespace WinTab.App.Services;

public static class AppEnvironment
{
    public static string ResolveLaunchExecutablePath()
    {
        string? processPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(processPath) && !IsDotNetHost(processPath))
            return processPath;

        string assemblyPath = Assembly.GetExecutingAssembly().Location;
        if (!string.IsNullOrWhiteSpace(assemblyPath))
        {
            string appHostPath = Path.ChangeExtension(assemblyPath, ".exe");
            if (File.Exists(appHostPath))
                return appHostPath;
        }

        return processPath ?? assemblyPath;
    }

    private static bool IsDotNetHost(string path)
    {
        string fileName = Path.GetFileName(path);
        return string.Equals(fileName, "dotnet", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(fileName, "dotnet.exe", StringComparison.OrdinalIgnoreCase);
    }

    public static void TryOpenFolderFallback(string path, WinTab.Diagnostics.Logger? logger)
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{path}\"",
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            logger?.Error($"Failed to fallback open-folder launch for path: {path}", ex);
        }
    }
}
