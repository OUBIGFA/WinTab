using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using AutoUpdaterDotNET;
using AutoUpdaterDotNET.Markdown;
using WinTab.Helpers;

namespace WinTab.Managers;

internal static class UpdateManager
{
    private static readonly HttpClient UpdateHttpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(10)
    };

    static UpdateManager()
    {
        UpdateHttpClient.DefaultRequestHeaders.UserAgent.ParseAdd("WinTab/1.0");
        AutoUpdater.FlattenRootFolder = true;
        using var icon = Helper.GetIcon();
        if (icon != null)
            AutoUpdater.Icon = icon.ToBitmap();
        AutoUpdater.ChangelogViewerProvider = new MarkdownViewerProvider();

        AutoUpdater.ParseUpdateInfoEvent += ParseUpdateInfo;
    }

    public static void CheckForUpdates() => AutoUpdater.Start(Constants.UpdateUrl);

    public static async Task<UpdateCheckResult> CheckForUpdatesWithResultAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            await using var stream = await UpdateHttpClient.GetStreamAsync(Constants.UpdateUrl, cancellationToken).ConfigureAwait(false);
            var jsonNode = await JsonSerializer.DeserializeAsync<JsonNode>(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
            if (jsonNode == null)
                return UpdateCheckResult.Failed();

            var tagName = jsonNode["tag_name"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(tagName) || !TryNormalizeVersion(tagName, out var latestVersion))
                return UpdateCheckResult.Failed();

            var installedVersion = NormalizeVersion(Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0));
            return new UpdateCheckResult(
                Completed: true,
                UpdateAvailable: latestVersion.CompareTo(installedVersion) > 0,
                LatestVersion: tagName,
                DownloadUrl: FindMatchingUpdateAssetUrl(jsonNode),
                ErrorMessage: null);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            Debug.WriteLine($"Manual update check failed: {ex.Message}");
            return UpdateCheckResult.Failed(ex.Message);
        }
    }

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

    private static bool TryNormalizeVersion(string value, out Version version)
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

    private static Version NormalizeVersion(Version version)
    {
        return new Version(
            Math.Max(0, version.Major),
            Math.Max(0, version.Minor),
            Math.Max(0, version.Build),
            Math.Max(0, version.Revision));
    }
}

internal sealed record UpdateCheckResult(
    bool Completed,
    bool UpdateAvailable,
    string? LatestVersion,
    string? DownloadUrl,
    string? ErrorMessage)
{
    public static UpdateCheckResult Failed(string? errorMessage = null) =>
        new(false, false, null, null, errorMessage);
}
