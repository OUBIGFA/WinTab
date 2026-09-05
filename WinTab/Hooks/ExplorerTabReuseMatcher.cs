using System;
using System.Collections.Generic;

namespace WinTab.Hooks;

internal sealed class ExplorerTabReuseCandidate
{
    private readonly Func<string?> _readLocation;
    private readonly Action<string?>? _updateLocation;

    public ExplorerTabReuseCandidate(
        nint tabHandle,
        string? cachedLocation,
        Func<string?> readLocation,
        Action<string?>? updateLocation = null)
    {
        TabHandle = tabHandle;
        CachedLocation = cachedLocation;
        _readLocation = readLocation;
        _updateLocation = updateLocation;
    }

    public nint TabHandle { get; }
    public string? CachedLocation { get; private set; }

    public string? ReadLocation() => _readLocation();

    public void UpdateLocation(string location)
    {
        CachedLocation = location;
        _updateLocation?.Invoke(location);
    }
}

internal static class ExplorerTabReuseMatcher
{
    public static bool TryFind(
        string targetLocation,
        IEnumerable<ExplorerTabReuseCandidate> candidates,
        Func<string, string, bool> areEquivalent,
        out nint tabHandle)
    {
        tabHandle = 0;
        var liveScanCandidates = new List<ExplorerTabReuseCandidate>();

        foreach (var candidate in candidates)
        {
            if (candidate.TabHandle == 0)
                continue;

            if (!HasCachedMatch(candidate, targetLocation, areEquivalent))
            {
                liveScanCandidates.Add(candidate);
                continue;
            }

            if (TryConfirmLiveMatch(candidate, targetLocation, areEquivalent))
            {
                tabHandle = candidate.TabHandle;
                return true;
            }
        }

        foreach (var candidate in liveScanCandidates)
        {
            if (!TryConfirmLiveMatch(candidate, targetLocation, areEquivalent))
                continue;

            tabHandle = candidate.TabHandle;
            return true;
        }

        return false;
    }

    private static bool HasCachedMatch(
        ExplorerTabReuseCandidate candidate,
        string targetLocation,
        Func<string, string, bool> areEquivalent)
    {
        if (string.IsNullOrWhiteSpace(candidate.CachedLocation))
            return false;

        try
        {
            return areEquivalent(targetLocation, candidate.CachedLocation);
        }
        catch
        {
            return false;
        }
    }

    private static bool TryConfirmLiveMatch(
        ExplorerTabReuseCandidate candidate,
        string targetLocation,
        Func<string, string, bool> areEquivalent)
    {
        string? liveLocation;
        try
        {
            liveLocation = candidate.ReadLocation();
        }
        catch
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(liveLocation))
            return false;

        try
        {
            candidate.UpdateLocation(liveLocation);
            return areEquivalent(targetLocation, liveLocation);
        }
        catch
        {
            return false;
        }
    }
}
