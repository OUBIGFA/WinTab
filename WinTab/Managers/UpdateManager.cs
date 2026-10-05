using System;
using System.Diagnostics;
using System.IO;
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
using WinTab.UI.Localization;

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

    public static Task<UpdateCheckResult> CheckForUpdatesWithResultAsync(CancellationToken cancellationToken = default) =>
        CheckForUpdatesWithResultAsync(UpdateHttpClient, cancellationToken);

    internal static async Task<UpdateCheckResult> CheckForUpdatesWithResultAsync(HttpClient client, CancellationToken cancellationToken = default)
    {
        using var requestCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        requestCancellation.CancelAfter(client.Timeout);
        try
        {
            await using var stream = await client.GetStreamAsync(Constants.UpdateUrl, requestCancellation.Token).ConfigureAwait(false);
            var jsonNode = await JsonSerializer.DeserializeAsync<JsonNode>(stream, cancellationToken: requestCancellation.Token).ConfigureAwait(false);
            if (!UpdateReleaseParser.TryParseRelease(jsonNode, RuntimeInformation.ProcessArchitecture, out var release))
                return UpdateCheckResult.Failed();

            var installedVersion = UpdateReleaseParser.NormalizeVersion(Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0));
            return new UpdateCheckResult(
                Completed: true,
                UpdateAvailable: release.Version.CompareTo(installedVersion) > 0,
                LatestVersion: release.TagName,
                DownloadUrl: release.DownloadUrl,
                ErrorMessage: null);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return UpdateCheckResult.Failed("The update request timed out. Please try again.");
        }
        catch (Exception ex) when (ex is HttpRequestException or IOException or JsonException or InvalidOperationException or FormatException)
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
            if (!UpdateReleaseParser.TryParseRelease(jsonNode, RuntimeInformation.ProcessArchitecture, out var release))
                return;

            p.UpdateInfo = new UpdateInfoEventArgs
            {
                CurrentVersion = release.DisplayVersion,
                ChangelogText = release.Changelog,
                ChangelogURL = release.ReleaseUrl,
                DownloadURL = release.DownloadUrl,
                // Keep setup in the chosen language and resume in the tray after an update.
                InstallerArgs = "/UPDATE=1 /LANG=" + (UiStrings.IsChinese ? "chinesesimplified" : "english"),
                // The download is checked against the digest GitHub published before the installer is run.
                CheckSum = release.DownloadSha256 is { } hash ? new CheckSum { Value = hash, HashingAlgorithm = "SHA256" } : null
            };
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Update check failed: {ex.Message}");
        }
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
