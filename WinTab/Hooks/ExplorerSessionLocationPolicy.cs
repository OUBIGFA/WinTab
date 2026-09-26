using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Models;

namespace WinTab.Hooks;

/// <summary>
/// Automatic restore must not ask an arbitrary shell extension or an offline share to resolve a PIDL.
/// Only Explorer's built-in local pages and existing directories on local drives are admitted; a junction
/// or symbolic link is followed only when it leads to another local drive. The filesystem probe runs off
/// the shell/UI threads; paths it has not confirmed within its budget are skipped, and at most one
/// abandoned probe can remain blocked.
/// </summary>
internal sealed class ExplorerSessionLocationPolicy
{
    private const int MaxLinkDepth = 8;
    private const string ThisPc = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
    private const string Home = "shell:::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}";
    private const string QuickAccess = "shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}";
    private const string RecycleBin = "shell:::{645FF040-5081-101B-9F08-00AA002F954E}";
    private const string Gallery = "shell:::{E88865EA-0E1C-4E20-9AA6-EDCD0212C87C}";
    private readonly Func<string, bool> _directoryAvailable;
    private readonly object _gate = new();
    private Task? _probe;

    public ExplorerSessionLocationPolicy(Func<string, bool>? directoryAvailable = null) =>
        _directoryAvailable = directoryAvailable ?? IsLocalDirectoryAvailable;

    /// <summary>Explorer's own start pages, shown by a plain launch unless the user chose another folder.</summary>
    internal static bool IsStartPage(string location) => IsOneOf(location, ThisPc, Home, QuickAccess);

    /// <summary>Built-in local pages that are restored without a filesystem or shell-extension probe.</summary>
    internal static bool IsKnownShellPage(string location) => IsOneOf(location, ThisPc, Home, QuickAccess, RecycleBin, Gallery);

    private static bool IsOneOf(string location, params string[] pages)
    {
        var normalized = Helper.NormalizeLocation(location);
        return Array.Exists(pages, page => normalized.Equals(page, StringComparison.OrdinalIgnoreCase));
    }

    internal static bool IsLocalPath(string location) => location.Length >= 3 &&
        char.IsAsciiLetter(location[0]) && location[1] == ':' && location[2] == '\\' &&
        location.IndexOfAny(['\0', '*', '?']) < 0 && location.AsSpan(2).IndexOf(':') < 0;

    public async Task<HashSet<int>> FindAvailableAsync(ExplorerSession session, CancellationToken cancellationToken,
        int timeoutMs = 1_500)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var locations = (string[])session.Locations.Clone();
        var available = new HashSet<int>();
        for (var index = 0; index < locations.Length; index++)
            if (IsKnownShellPage(locations[index]))
                available.Add(index);

        var confirmed = new bool[locations.Length];
        Task probe;
        lock (_gate)
        {
            // A timeout does not stop a synchronous filesystem call. Do not queue more calls behind it
            // or let repeated launches exhaust the thread pool while a device is unavailable.
            if (_probe is { IsCompleted: false })
                return available;
            _probe = probe = Task.Run(() => Probe(locations, confirmed, cancellationToken), cancellationToken);
        }
        try
        {
            await probe.WaitAsync(TimeSpan.FromMilliseconds(timeoutMs), cancellationToken).ConfigureAwait(false);
        }
        catch (TimeoutException)
        {
            ExplorerDebugLog.Write("Session restore filesystem check timed out; unconfirmed paths were skipped.");
        }
        // Paths confirmed before the budget ended stay restorable; a blocked path skips only itself and later ones.
        for (var index = 0; index < confirmed.Length; index++)
            if (Volatile.Read(ref confirmed[index]))
                available.Add(index);
        return available;
    }

    private void Probe(string[] locations, bool[] confirmed, CancellationToken cancellationToken)
    {
        for (var index = 0; index < locations.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var location = Helper.NormalizeLocation(locations[index]);
            if (!IsKnownShellPage(location) && IsLocalPath(location) && _directoryAvailable(location))
                Volatile.Write(ref confirmed[index], true);
        }
    }

    private static bool IsLocalDirectoryAvailable(string location)
    {
        try { return IsLocalDirectory(Path.GetFullPath(location), 0); }
        catch (Exception exception) when (IsProbeFailure(exception)) { return false; }
    }

    /// <summary>
    /// Checks each ancestor before relying on it: a network-backed link must never be followed. A reparse
    /// point that is not a link (for example a cloud-synced folder) stays at its local path.
    /// </summary>
    private static bool IsLocalDirectory(string fullPath, int depth)
    {
        var root = Path.GetPathRoot(fullPath);
        if (depth > MaxLinkDepth || root == null || !IsLocalPath(root))
            return false;
        var drive = new DriveInfo(root);
        if (drive.DriveType is not (DriveType.Fixed or DriveType.Removable) || !drive.IsReady)
            return false;

        var current = root;
        if (!IsUsableDirectory(current, depth))
            return false;
        foreach (var segment in fullPath[root.Length..].Split('\\', StringSplitOptions.RemoveEmptyEntries))
        {
            current = Path.Combine(current, segment);
            if (!IsUsableDirectory(current, depth))
                return false;
        }
        return true;
    }

    private static bool IsUsableDirectory(string path, int depth)
    {
        var attributes = File.GetAttributes(path);
        if (!attributes.HasFlag(FileAttributes.Directory) || attributes.HasFlag(FileAttributes.Offline))
            return false;
        if (!attributes.HasFlag(FileAttributes.ReparsePoint))
            return true;
        var target = new DirectoryInfo(path).LinkTarget;
        if (target == null)
            return true;
        var resolved = ResolveLinkTarget(path, target);
        return resolved != null && IsLocalDirectory(resolved, depth + 1);
    }

    /// <summary>The drive-letter path a link leads to; null for shares, volume GUIDs and process-relative forms.</summary>
    internal static string? ResolveLinkTarget(string linkPath, string target)
    {
        string resolved;
        if (target.StartsWith(@"\??\", StringComparison.Ordinal) || target.StartsWith(@"\\?\", StringComparison.Ordinal))
            resolved = target[4..];
        else if (Path.IsPathFullyQualified(target))
            resolved = target;
        else if (!Path.IsPathRooted(target))
            resolved = Path.Combine(Path.GetDirectoryName(linkPath) ?? string.Empty, target);
        else
            return null;
        return IsLocalPath(resolved) ? Path.GetFullPath(resolved) : null;
    }

    private static bool IsProbeFailure(Exception exception) => exception is IOException or UnauthorizedAccessException or
        ArgumentException or NotSupportedException or System.Security.SecurityException;
}
