using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using AutoUpdaterDotNET;
using AutoUpdaterDotNET.Markdown;
using WinTab.Helpers;

namespace WinTab.Managers;

internal static class UpdateManager
{
    static UpdateManager()
    {
        AutoUpdater.FlattenRootFolder = true;
        using var icon = Helper.GetIcon();
        if (icon != null)
            AutoUpdater.Icon = icon.ToBitmap();
        AutoUpdater.ChangelogViewerProvider = new MarkdownViewerProvider();

        AutoUpdater.ParseUpdateInfoEvent += ParseUpdateInfo;
    }

    public static void CheckForUpdates() => AutoUpdater.Start(Constants.UpdateUrl);

    private static void ParseUpdateInfo(ParseUpdateInfoEventArgs p)
    {
        try
        {
            var jsonNode = JsonSerializer.Deserialize<JsonNode>(p.RemoteData);
            if (jsonNode == null) return;

            var tagName = jsonNode["tag_name"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(tagName)) return;

            p.UpdateInfo = new UpdateInfoEventArgs
            {
                CurrentVersion = tagName.TrimStart('v'),
                ChangelogText = jsonNode["body"]?.GetValue<string>() ?? string.Empty,
                ChangelogURL = jsonNode["html_url"]?.GetValue<string>() ?? string.Empty,
                DownloadURL = FindMatchingUpdateAssetUrl(jsonNode)
            };
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Update check failed: {ex.Message}");
        }
    }

    private static string? FindMatchingUpdateAssetUrl(JsonNode jsonNode)
    {
        if (jsonNode["assets"] is not JsonArray assets)
            return null;

        var setupAssets = new List<(string Name, string Url)>();
        foreach (var asset in assets)
        {
            var assetName = asset?["name"]?.GetValue<string>();
            var downloadUrl = asset?["browser_download_url"]?.GetValue<string>();

            if (!string.IsNullOrWhiteSpace(assetName) &&
                !string.IsNullOrWhiteSpace(downloadUrl) &&
                assetName.EndsWith("_Setup.exe", StringComparison.OrdinalIgnoreCase))
            {
                setupAssets.Add((assetName, downloadUrl));
            }
        }

        if (setupAssets.Count == 0)
            return null;

        var architectureSuffix = GetInstallerArchitectureSuffix();
        if (architectureSuffix != null)
        {
            foreach (var asset in setupAssets)
            {
                if (asset.Name.EndsWith(architectureSuffix, StringComparison.OrdinalIgnoreCase))
                    return asset.Url;
            }
        }

        return setupAssets[0].Url;
    }

    private static string? GetInstallerArchitectureSuffix()
    {
        return RuntimeInformation.ProcessArchitecture switch
        {
            Architecture.X64 => "_x64_Setup.exe",
            Architecture.X86 => "_x86_Setup.exe",
            Architecture.Arm64 => "_arm64_Setup.exe",
            _ => null
        };
    }
}
