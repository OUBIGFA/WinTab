using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;

namespace WinTab.Managers;

/// <param name="TagName">The release tag as published, such as v2.1.0.</param>
/// <param name="Version">The tag as a version with every component present, comparable with the installed one.</param>
/// <param name="DisplayVersion">The version as shown to the user, without an unused fourth component.</param>
/// <param name="DownloadUrl">The installer for this process's architecture, or null when the release has none.</param>
internal sealed record UpdateRelease(string TagName, Version Version, string DisplayVersion, string Changelog,
    string ReleaseUrl, string? DownloadUrl);

/// <summary>
/// Reads GitHub's latest-release answer. The manual check and the automatic updater both read it here, so they
/// agree on which tags are versions and which installer belongs to this process.
/// </summary>
internal static class UpdateReleaseParser
{
    /// <summary>False when the answer is not a release object or its tag is not a version.</summary>
    public static bool TryParseRelease(JsonNode? releaseNode, Architecture architecture, out UpdateRelease release)
    {
        release = null!;
        if (releaseNode is not JsonObject releaseObject ||
            ReadString(releaseObject, "tag_name") is not { } tagName || string.IsNullOrWhiteSpace(tagName) ||
            !TryNormalizeVersion(tagName, out var version))
            return false;

        release = new UpdateRelease(tagName, version,
            version.Revision == 0 ? version.ToString(3) : version.ToString(),
            ReadString(releaseObject, "body") ?? string.Empty,
            ReadString(releaseObject, "html_url") ?? string.Empty,
            FindMatchingAssetUrl(releaseObject, architecture));
        return true;
    }

    public static string? FindMatchingAssetUrl(JsonNode releaseNode, Architecture architecture)
    {
        if (releaseNode is not JsonObject releaseObject || releaseObject["assets"] is not JsonArray assets)
            return null;

        var setupAssets = new List<(string Name, string Url)>();
        foreach (var asset in assets)
        {
            if (asset is not JsonObject assetObject)
                continue;
            var assetName = ReadString(assetObject, "name");
            var downloadUrl = ReadString(assetObject, "browser_download_url");

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

    /// <summary>A text property, or null when it is missing or of another type.</summary>
    private static string? ReadString(JsonObject node, string name) =>
        node[name] is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;

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
