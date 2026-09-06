using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;

namespace WinTab.Managers;

internal static class UpdateReleaseParser
{
    public static string? FindMatchingAssetUrl(JsonNode releaseNode, Architecture architecture)
    {
        if (releaseNode["assets"] is not JsonArray assets)
            return null;

        var setupAssets = new List<(string Name, string Url)>();
        foreach (var asset in assets)
        {
            var assetName = asset?["name"]?.GetValue<string>();
            var downloadUrl = asset?["browser_download_url"]?.GetValue<string>();

            if (!string.IsNullOrWhiteSpace(assetName) &&
                !string.IsNullOrWhiteSpace(downloadUrl) &&
                Uri.TryCreate(downloadUrl, UriKind.Absolute, out var downloadUri) &&
                downloadUri.Scheme == Uri.UriSchemeHttps &&
                assetName.EndsWith("_Setup.exe", StringComparison.OrdinalIgnoreCase))
            {
                setupAssets.Add((assetName, downloadUrl));
            }
        }

        if (setupAssets.Count == 0)
            return null;

        var architectureSuffix = GetInstallerArchitectureSuffix(architecture);
        if (architectureSuffix != null)
        {
            foreach (var asset in setupAssets)
            {
                if (asset.Name.EndsWith(architectureSuffix, StringComparison.OrdinalIgnoreCase))
                    return asset.Url;
            }
        }

        return null;
    }

    public static bool TryNormalizeVersion(string value, out Version version)
    {
        var normalized = value.Trim().TrimStart('v', 'V');
        var suffixIndex = normalized.IndexOfAny(['-', '+']);
        if (suffixIndex >= 0)
            normalized = normalized[..suffixIndex];

        if (!Version.TryParse(normalized, out var parsed))
        {
            version = new Version(0, 0);
            return false;
        }

        version = NormalizeVersion(parsed);
        return true;
    }

    public static Version NormalizeVersion(Version version)
    {
        return new Version(
            Math.Max(0, version.Major),
            Math.Max(0, version.Minor),
            Math.Max(0, version.Build),
            Math.Max(0, version.Revision));
    }

    private static string? GetInstallerArchitectureSuffix(Architecture architecture)
    {
        return architecture switch
        {
            Architecture.X64 => "_x64_Setup.exe",
            Architecture.X86 => "_x86_Setup.exe",
            Architecture.Arm64 => "_arm64_Setup.exe",
            _ => null
        };
    }
}
