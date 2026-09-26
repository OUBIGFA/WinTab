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
    /// <summary>UTC ticks of the save; 0 in files written before the live journal existed.</summary>
    public long SavedAt { get; init; }
    /// <summary>Set only on the live journal: the Explorer process that still showed the window.</summary>
    public ExplorerSessionOwner? Owner { get; init; }

    public ExplorerSession Copy() => this with { Locations = (string[])Locations.Clone() };

    /// <summary>The same tabs in the same state, whenever and by whichever process they were recorded.</summary>
    public bool HasSameTabs(ExplorerSession? other) => other != null &&
        Locations.SequenceEqual(other.Locations, StringComparer.Ordinal) &&
        ActiveTabIndex == other.ActiveTabIndex && OrderVerified == other.OrderVerified;

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
        if (SavedAt < 0 || SavedAt > DateTime.MaxValue.Ticks ||
            Owner is { } owner && (owner.ProcessId <= 0 || owner.StartedAt <= 0 || owner.StartedAt > DateTime.MaxValue.Ticks))
            throw new JsonException("The session save time or process owner is invalid.");
        return Copy();
    }

    internal sealed class UnsupportedVersionException(int version)
        : JsonException($"Session format {version} is not supported; the file will not be overwritten.");
}

/// <param name="StartedAt">UTC ticks of the process start; a reused process id never matches it.</param>
internal sealed record ExplorerSessionOwner(int ProcessId, long StartedAt)
{
    public static ExplorerSessionOwner? Of(int processId)
    {
        try
        {
            using var process = System.Diagnostics.Process.GetProcessById(processId);
            return new ExplorerSessionOwner(processId, process.StartTime.ToUniversalTime().Ticks);
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException or
            System.ComponentModel.Win32Exception or NotSupportedException)
        {
            return null;
        }
    }

    /// <summary>
    /// False only when the owner is known to be gone: the id is unused or now names a process started at
    /// another time. An owner that cannot be inspected is assumed to still show its windows.
    /// </summary>
    public bool IsRunning()
    {
        try
        {
            using var process = System.Diagnostics.Process.GetProcessById(ProcessId);
            return !process.HasExited && process.StartTime.ToUniversalTime().Ticks == StartedAt;
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException)
        {
            return false;
        }
        catch (Exception exception) when (exception is System.ComponentModel.Win32Exception or NotSupportedException)
        {
            return true;
        }
    }
}
