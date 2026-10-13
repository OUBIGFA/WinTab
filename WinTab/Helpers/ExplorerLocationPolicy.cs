using System;

namespace WinTab.Helpers;

/// <summary>
/// Control Panel shares Explorer's window classes but is not a folder-tab host. Use its shell namespace,
/// not localized titles or filesystem names, to keep merging, reuse and restoration out of that window.
/// </summary>
internal static class ExplorerLocationPolicy
{
    private static readonly string[] ControlPanelRoots =
    [
        "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", // Category view, including \\0 (all items).
        "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}", // Legacy all-items namespace and its child pages.
        "shell:::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}", // Start menu / desktop Control Panel command object.
        "shell:ControlPanelFolder"
    ];

    internal static bool IsControlPanel(string? location)
    {
        if (string.IsNullOrWhiteSpace(location))
            return false;
        var normalized = Helper.NormalizeLocation(location);
        foreach (var root in ControlPanelRoots)
        {
            if (normalized.Equals(root, StringComparison.OrdinalIgnoreCase) ||
                normalized.StartsWith(root, StringComparison.OrdinalIgnoreCase) &&
                normalized.Length > root.Length && normalized[root.Length] == '\\')
                return true;
        }
        return false;
    }
}
