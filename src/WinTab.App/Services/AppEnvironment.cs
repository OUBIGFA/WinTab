using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using WinTab.Platform.Win32;

namespace WinTab.App.Services;

public static class AppEnvironment
{
    private const int MaxOpenPathLength = 2048;
    internal static Func<string, bool> TryOpenNativeShellTarget = static target => NativeShellLauncher.TryOpen(target);
    internal static Func<string, bool> StartExplorerProcess = static normalizedPath =>
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"\"{normalizedPath}\"",
            UseShellExecute = true,
        });
        return true;
    };

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
        _ = TryOpenTargetFallback(path, logger);
    }

    public static bool TryOpenTargetFallback(string path, WinTab.Diagnostics.Logger? logger)
    {
        if (!TryNormalizeExistingDirectoryPath(path, out string normalizedPath, out string reason))
        {
            logger?.Warn($"Skipped Explorer fallback launch: invalid path ({reason}). Raw='{path}'");
            return false;
        }

        OpenTargetInfo targetInfo = OpenTargetClassifier.Classify(normalizedPath);
        if (targetInfo.RequiresNativeShellLaunch)
        {
            bool openedNatively = TryOpenNativeShellTarget(normalizedPath);
            if (!openedNatively)
                logger?.Warn($"Failed native shell launch for target: {normalizedPath}");

            return openedNatively;
        }

        try
        {
            return StartExplorerProcess(normalizedPath);
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or FileNotFoundException)
        {
            logger?.Error($"Failed to fallback open-folder launch for path: {normalizedPath}", ex);
            return false;
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

        // Keep shell namespace support (e.g. "::{GUID}", "{GUID}", "shell:RecycleBinFolder").
        if (TryNormalizeShellNamespaceToken(trimmedPath, out string namespaceToken))
        {
            normalizedPath = namespaceToken;
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

    private static bool TryNormalizeShellNamespaceToken(string value, out string normalized)
    {
        return ShellNamespacePath.TryNormalizeToken(value, out normalized);
    }
}
