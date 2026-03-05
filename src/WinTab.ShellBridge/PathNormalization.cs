using System.IO;

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

        if (trimmedPath.StartsWith("::", StringComparison.Ordinal) ||
            trimmedPath.StartsWith("shell:", StringComparison.OrdinalIgnoreCase))
        {
            normalizedPath = trimmedPath;
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
}
