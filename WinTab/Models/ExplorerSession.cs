using System;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace WinTab.Models;

/// <summary>One closed Explorer window, including duplicate paths and a single-tab window.</summary>
internal sealed record ExplorerSession
{
    public const int CurrentVersion = 1;
    public const int MaxTabs = 100;
    private const int MaxLocationLength = 32_767;
    private const int MaxTotalLocationLength = 512 * 1024;

    [JsonRequired]
    public int Version { get; init; } = CurrentVersion;
    [JsonRequired]
    public string[] Locations { get; init; } = [];
    [JsonRequired]
    public int ActiveTabIndex { get; init; }
    [JsonRequired]
    public bool OrderVerified { get; init; }

    public ExplorerSession Copy() => this with { Locations = (string[])Locations.Clone() };

    public ExplorerSession ValidatedCopy()
    {
        if (Version != CurrentVersion)
            throw new UnsupportedVersionException(Version);
        if (Locations == null || Locations.Length is 0 or > MaxTabs ||
            ActiveTabIndex < 0 || ActiveTabIndex >= Locations.Length)
            throw new JsonException("The session must contain 1 to 100 tabs and a valid active-tab index.");
        if (Locations.Any(location => string.IsNullOrWhiteSpace(location) || location.Length > MaxLocationLength ||
            location.Any(char.IsControl)) || Locations.Sum(location => (long)location.Length) > MaxTotalLocationLength)
            throw new JsonException("The session contains an invalid or oversized location.");
        return Copy();
    }

    internal sealed class UnsupportedVersionException(int version)
        : JsonException($"Session format {version} is not supported; the file will not be overwritten.");
}
