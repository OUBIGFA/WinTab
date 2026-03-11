namespace WinTab.Platform.Win32;

public static class OpenTargetClassifier
{
    private static readonly ShellLocationIdentityService LocationIdentity = new();

    public static OpenTargetInfo Classify(string? target)
    {
        string rawTarget = target?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(rawTarget))
            return new OpenTargetInfo(rawTarget, string.Empty, OpenTargetKind.Invalid);

        string normalizedTarget = LocationIdentity.NormalizeLocation(rawTarget);
        if (string.IsNullOrWhiteSpace(normalizedTarget))
            return new OpenTargetInfo(rawTarget, normalizedTarget, OpenTargetKind.Invalid);

        if (IsPhysicalFileSystemPath(normalizedTarget))
            return new OpenTargetInfo(rawTarget, normalizedTarget, OpenTargetKind.PhysicalFileSystem);

        if (ShellNamespacePath.IsShellNamespace(normalizedTarget))
            return new OpenTargetInfo(rawTarget, normalizedTarget, OpenTargetKind.ShellNamespace);

        return new OpenTargetInfo(rawTarget, normalizedTarget, OpenTargetKind.Invalid);
    }

    public static bool IsPhysicalFileSystemPath(string? target)
    {
        if (string.IsNullOrWhiteSpace(target))
            return false;

        string value = target.Trim();
        if (value.StartsWith("\\\\", StringComparison.OrdinalIgnoreCase))
            return true;

        return value.Length >= 3 &&
               char.IsLetter(value[0]) &&
               value[1] == ':' &&
               (value[2] == '\\' || value[2] == '/');
    }
}
