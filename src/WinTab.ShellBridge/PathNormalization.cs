using System;
using System.IO;
using WinTab.Platform.Win32;

namespace WinTab.ShellBridge;

internal static class PathNormalization
{
    private const int MaxOpenPathLength = 2048;

    public static bool TryNormalizeOpenTarget(string? candidatePath, out string normalizedPath)
    {
        normalizedPath = string.Empty;

        if (string.IsNullOrWhiteSpace(candidatePath))
            return false;

        string trimmedPath = candidatePath.Trim().Trim('"');
        if (trimmedPath.Length == 0 || trimmedPath.Length > MaxOpenPathLength)
            return false;

        foreach (char character in trimmedPath)
        {
            if (char.IsControl(character))
                return false;
        }

        if (TryNormalizeShellNamespaceToken(trimmedPath, out string namespaceToken))
        {
            normalizedPath = namespaceToken;
            return true;
        }

        if (trimmedPath.Length == 2 &&
            char.IsLetter(trimmedPath[0]) &&
            trimmedPath[1] == ':')
        {
            normalizedPath = $"{trimmedPath}\\";
            return true;
        }

        try
        {
            normalizedPath = Path.GetFullPath(trimmedPath);
            return Path.IsPathRooted(normalizedPath);
        }
        catch
        {
            return false;
        }
    }

    private static bool TryNormalizeShellNamespaceToken(string value, out string normalized)
    {
        return ShellNamespacePath.TryNormalizeToken(value, out normalized);
    }
}
