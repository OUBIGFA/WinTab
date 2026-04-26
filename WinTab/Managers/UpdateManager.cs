using System;
using System.Diagnostics;
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
        AutoUpdater.Icon = Helper.GetIcon()!.ToBitmap();
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

            p.UpdateInfo = new UpdateInfoEventArgs
            {
                CurrentVersion = jsonNode["tag_name"]!.GetValue<string>().TrimStart('v'),
                ChangelogText = jsonNode["body"]!.GetValue<string>(),
                ChangelogURL = jsonNode["html_url"]!.GetValue<string>(),
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
        var assets = jsonNode["assets"]!.AsArray();
        foreach (var asset in assets)
        {
            var assetName = asset!["name"]!.GetValue<string>();
            if (assetName.EndsWith("_Setup.exe", StringComparison.OrdinalIgnoreCase))
                return asset["browser_download_url"]!.GetValue<string>();
        }

        return null;
    }
}
