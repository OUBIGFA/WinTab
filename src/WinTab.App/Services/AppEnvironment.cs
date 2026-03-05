using System;
using System.IO;
using System.Reflection;

namespace WinTab.App.Services;

public static class AppEnvironment
{
    private const int MaxOpenPathLength = 2048;

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
        if (!TryNormalizeExistingDirectoryPath(path, out string normalizedPath, out string reason))
        {
            logger?.Warn($"Skipped Explorer fallback launch: invalid path ({reason}). Raw='{path}'");
            return;
        }

        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{normalizedPath}\"",
                UseShellExecute = true,
            });
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or FileNotFoundException)
        {
            logger?.Error($"Failed to fallback open-folder launch for path: {normalizedPath}", ex);
        }
    }

    public static bool TryNormalizeExistingDirectoryPath(string? candidatePath, out string normalizedPath, out string failureReason)
    {
        normalizedPath = string.Empty;
        failureReason = string.Empty;

        if (string.IsNullOrWhiteSpace(candidatePath))
        {
            failureReason = "empty path";
            return false;
        }

        string trimmedPath = candidatePath.Trim().Trim('"');
        if (trimmedPath.Length == 0)
        {
            failureReason = "empty path";
            return false;
        }

        if (trimmedPath.Length > MaxOpenPathLength)
        {
            failureReason = "path too long";
            return false;
        }

        if (trimmedPath.Contains("://", StringComparison.Ordinal))
        {
            failureReason = "unsupported URI scheme";
            return false;
        }

        foreach (char character in trimmedPath)
        {
            if (char.IsControl(character))
            {
                failureReason = "contains control characters";
                return false;
            }
        }

        // Keep shell namespace support (e.g. "::{GUID}") and permissive Explorer-open semantics.
        // Only normalize rooted file-system paths; for shell targets keep original token.
        if (trimmedPath.StartsWith("::", StringComparison.Ordinal) ||
            trimmedPath.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
        {
            normalizedPath = trimmedPath;
            return true;
        }

        // Drive designator-only input (e.g. "C:") from shell invocations can arrive as "C:".
        // Do not resolve it against process current directory, always normalize to drive root.
        if (trimmedPath.Length == 2 &&
            char.IsLetter(trimmedPath[0]) &&
            trimmedPath[1] == ':')
        {
            normalizedPath = $"{trimmedPath}\\";
            return true;
        }

        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(trimmedPath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            failureReason = $"invalid path format ({ex.GetType().Name})";
            return false;
        }

        if (!Path.IsPathRooted(fullPath))
        {
            failureReason = "path is not rooted";
            return false;
        }

        normalizedPath = fullPath;
        return true;
    }
}
