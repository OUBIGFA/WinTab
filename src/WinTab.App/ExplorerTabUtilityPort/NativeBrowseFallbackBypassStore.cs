using System.Collections.Concurrent;
using WinTab.Platform.Win32;

namespace WinTab.App.ExplorerTabUtilityPort;

/// <summary>
/// Tracks short-lived location tokens that indicate a path should be opened as a native Explorer window.
/// Matching windows must not be auto-converted into tabs.
/// </summary>
internal sealed class NativeBrowseFallbackBypassStore
{
    private readonly ShellLocationIdentityService _locationIdentity;
    private readonly ConcurrentDictionary<string, DateTimeOffset> _locationBypassUntilUtc =
        new(StringComparer.OrdinalIgnoreCase);

    public NativeBrowseFallbackBypassStore(ShellLocationIdentityService locationIdentity)
    {
        ArgumentNullException.ThrowIfNull(locationIdentity);
        _locationIdentity = locationIdentity;
    }

    public void Register(string location, TimeSpan ttl)
    {
        CleanupExpired();

        if (!TryNormalizeKey(location, out string key))
            return;

        if (ttl <= TimeSpan.Zero)
        {
            _locationBypassUntilUtc.TryRemove(key, out _);
            return;
        }

        _locationBypassUntilUtc[key] = DateTimeOffset.UtcNow.Add(ttl);
    }

    public bool TryConsume(string location)
    {
        if (!TryNormalizeKey(location, out string key))
            return false;

        if (!_locationBypassUntilUtc.TryGetValue(key, out DateTimeOffset untilUtc))
            return false;

        if (untilUtc <= DateTimeOffset.UtcNow)
        {
            _locationBypassUntilUtc.TryRemove(key, out _);
            return false;
        }

        return _locationBypassUntilUtc.TryRemove(key, out _);
    }

    public void Revoke(string location)
    {
        if (!TryNormalizeKey(location, out string key))
            return;

        _locationBypassUntilUtc.TryRemove(key, out _);
    }

    public void CleanupExpired()
    {
        DateTimeOffset now = DateTimeOffset.UtcNow;
        foreach (var kvp in _locationBypassUntilUtc)
        {
            if (kvp.Value <= now)
                _locationBypassUntilUtc.TryRemove(kvp.Key, out _);
        }
    }

    private bool TryNormalizeKey(string? location, out string key)
    {
        key = string.Empty;

        if (string.IsNullOrWhiteSpace(location))
            return false;

        key = _locationIdentity.NormalizeLocation(location);
        return !string.IsNullOrWhiteSpace(key);
    }
}
