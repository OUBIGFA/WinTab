using System.Security.Cryptography;
using System.IO;
using System.Text;
using System.Text.Json;
using Microsoft.Win32;
using WinTab.Diagnostics;
using WinTab.Persistence;

namespace WinTab.App.Services;

internal sealed class ExplorerShellBaselineStore
{
    private const string BackupRegistryPath = @"Software\WinTab\Backups\ExplorerOpenVerb";
    private const string BackupRegistryValueName = "BackupJson";
    private static readonly string[] TargetClasses =
    [
        "Folder",
        "Directory",
        "Drive"
    ];

    private readonly string _exePath;
    private readonly Logger _logger;

    public ExplorerShellBaselineStore(string exePath, Logger logger)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(exePath);
        ArgumentNullException.ThrowIfNull(logger);

        _exePath = exePath;
        _logger = logger;
    }

    public static string BackupPath => Path.Combine(AppPaths.BaseDirectory, "explorer-open-verb-backup.json");

    public void EnsureBaselineExists()
    {
        if (TryLoadBaseline() is not null)
            return;

        CaptureAndPersist();
    }

    public void CaptureAndPersist()
    {
        Directory.CreateDirectory(AppPaths.BaseDirectory);

        ExplorerShellBaseline baseline = CaptureBaseline();
        string json = Serialize(baseline);
        File.WriteAllText(BackupPath, json);

        using RegistryKey key = Registry.CurrentUser.CreateSubKey(BackupRegistryPath, writable: true)
            ?? throw new InvalidOperationException("Failed to create backup registry key.");
        key.SetValue(BackupRegistryValueName, json, RegistryValueKind.String);

        _logger.Info($"Captured Explorer shell baseline: {BackupPath}");
    }

    public ExplorerShellBaseline? TryLoadBaseline()
    {
        try
        {
            if (File.Exists(BackupPath))
            {
                string fileJson = File.ReadAllText(BackupPath);
                ExplorerShellBaseline? fileBaseline = DeserializeAndValidate(fileJson);
                if (fileBaseline is not null)
                    return fileBaseline;
            }
        }
        catch
        {
            // ignore and fall back to registry copy
        }

        try
        {
            using RegistryKey? key = Registry.CurrentUser.OpenSubKey(BackupRegistryPath, writable: false);
            string? json = key?.GetValue(BackupRegistryValueName) as string;
            if (string.IsNullOrWhiteSpace(json))
                return null;

            return DeserializeAndValidate(json);
        }
        catch
        {
            return null;
        }
    }

    public void DeletePersistedBaseline()
    {
        try
        {
            if (File.Exists(BackupPath))
                File.Delete(BackupPath);
        }
        catch
        {
            // ignore
        }

        try
        {
            using RegistryKey? root = Registry.CurrentUser.OpenSubKey(@"Software", writable: true);
            root?.DeleteSubKeyTree(@"WinTab\Backups\ExplorerOpenVerb", throwOnMissingSubKey: false);
        }
        catch
        {
            // ignore
        }
    }

    public void RestoreCurrentUserBaseline(ExplorerShellBaseline baseline)
    {
        ArgumentNullException.ThrowIfNull(baseline);

        foreach (RegistryTreeBackup entry in baseline.CurrentUserShellEntries)
            entry.Restore();
    }

    private ExplorerShellBaseline CaptureBaseline()
    {
        var baseline = new ExplorerShellBaseline
        {
            ExePath = _exePath,
            CreatedAtUtc = DateTimeOffset.UtcNow,
            CurrentUserShellEntries = CaptureShellEntries(RegistryHive.CurrentUser),
            LocalMachineShellEntries = CaptureShellEntries(RegistryHive.LocalMachine),
        };

        baseline.Sha256 = ComputeSha256(Serialize(baseline with { Sha256 = null }));
        return baseline;
    }

    private static List<RegistryTreeBackup> CaptureShellEntries(RegistryHive hive)
    {
        var entries = new List<RegistryTreeBackup>();
        foreach (RegistryView view in GetRegistryViews())
        {
            foreach (string cls in TargetClasses)
            {
                entries.Add(RegistryTreeBackup.Capture(hive, view, $@"Software\Classes\{cls}\shell"));
            }
        }

        return entries;
    }

    private static RegistryView[] GetRegistryViews()
    {
        return Environment.Is64BitOperatingSystem
            ? [RegistryView.Registry64, RegistryView.Registry32]
            : [RegistryView.Registry32];
    }

    private static string Serialize(ExplorerShellBaseline baseline)
    {
        return JsonSerializer.Serialize(baseline, new JsonSerializerOptions { WriteIndented = true });
    }

    private static ExplorerShellBaseline? DeserializeAndValidate(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
            return null;

        ExplorerShellBaseline? baseline = JsonSerializer.Deserialize<ExplorerShellBaseline>(json);
        if (baseline is null)
            return null;

        string expected = ComputeSha256(Serialize(baseline with { Sha256 = null }));
        return string.Equals(expected, baseline.Sha256, StringComparison.OrdinalIgnoreCase)
            ? baseline
            : null;
    }

    private static string ComputeSha256(string text)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        byte[] hash = SHA256.HashData(bytes);
        return Convert.ToHexString(hash);
    }
}

internal sealed record ExplorerShellBaseline
{
    public required string ExePath { get; init; }
    public required DateTimeOffset CreatedAtUtc { get; init; }
    public required List<RegistryTreeBackup> CurrentUserShellEntries { get; init; }
    public required List<RegistryTreeBackup> LocalMachineShellEntries { get; init; }
    public string? Sha256 { get; set; }
}

internal sealed record RegistryTreeBackup
{
    public required RegistryHive Hive { get; init; }
    public required RegistryView View { get; init; }
    public required string SubKeyPath { get; init; }
    public RegistryNodeBackup? Node { get; init; }

    public static RegistryTreeBackup Capture(RegistryHive hive, RegistryView view, string subKeyPath)
    {
        using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, view);
        using RegistryKey? key = baseKey.OpenSubKey(subKeyPath, writable: false);

        return new RegistryTreeBackup
        {
            Hive = hive,
            View = view,
            SubKeyPath = subKeyPath,
            Node = key is null ? null : RegistryNodeBackup.Capture(key),
        };
    }

    public void Restore()
    {
        using RegistryKey baseKey = RegistryKey.OpenBaseKey(Hive, View);
        baseKey.DeleteSubKeyTree(SubKeyPath, throwOnMissingSubKey: false);

        if (Node is null)
            return;

        using RegistryKey restoredKey = baseKey.CreateSubKey(SubKeyPath, writable: true)
            ?? throw new InvalidOperationException($"Failed to recreate registry key {Hive}\\{SubKeyPath} ({View}).");
        Node.RestoreInto(restoredKey);
    }
}

internal sealed record RegistryNodeBackup
{
    public required List<RegistryValueBackup> Values { get; init; }
    public required List<RegistrySubKeyBackup> SubKeys { get; init; }

    public static RegistryNodeBackup Capture(RegistryKey key)
    {
        var values = new List<RegistryValueBackup>();
        foreach (string valueName in key.GetValueNames())
        {
            RegistryValueBackup? value = RegistryValueBackup.FromRegistry(key, valueName);
            if (value is not null)
                values.Add(value);
        }

        var subKeys = new List<RegistrySubKeyBackup>();
        foreach (string subKeyName in key.GetSubKeyNames())
        {
            using RegistryKey subKey = key.OpenSubKey(subKeyName, writable: false)
                ?? throw new InvalidOperationException($"Failed to open registry subkey: {key.Name}\\{subKeyName}");
            subKeys.Add(new RegistrySubKeyBackup
            {
                Name = subKeyName,
                Node = Capture(subKey),
            });
        }

        return new RegistryNodeBackup
        {
            Values = values,
            SubKeys = subKeys,
        };
    }

    public void RestoreInto(RegistryKey key)
    {
        foreach (RegistryValueBackup value in Values)
            value.Apply(key);

        foreach (RegistrySubKeyBackup subKey in SubKeys)
        {
            using RegistryKey restoredSubKey = key.CreateSubKey(subKey.Name, writable: true)
                ?? throw new InvalidOperationException($"Failed to recreate registry subkey: {key.Name}\\{subKey.Name}");
            subKey.Node.RestoreInto(restoredSubKey);
        }
    }
}

internal sealed record RegistrySubKeyBackup
{
    public required string Name { get; init; }
    public required RegistryNodeBackup Node { get; init; }
}

internal sealed record RegistryValueBackup
{
    public required string Name { get; init; }
    public required RegistryValueKind Kind { get; init; }
    public string? StringValue { get; init; }
    public string[]? MultiStringValue { get; init; }
    public byte[]? BinaryValue { get; init; }
    public int? DwordValue { get; init; }
    public long? QwordValue { get; init; }

    public static RegistryValueBackup? FromRegistry(RegistryKey key, string valueName)
    {
        object? value = key.GetValue(valueName);
        if (value is null)
            return null;

        RegistryValueKind kind = key.GetValueKind(valueName);
        return kind switch
        {
            RegistryValueKind.String or RegistryValueKind.ExpandString =>
                new RegistryValueBackup { Name = valueName, Kind = kind, StringValue = value as string },
            RegistryValueKind.MultiString =>
                new RegistryValueBackup { Name = valueName, Kind = kind, MultiStringValue = value as string[] },
            RegistryValueKind.DWord =>
                new RegistryValueBackup { Name = valueName, Kind = kind, DwordValue = Convert.ToInt32(value) },
            RegistryValueKind.QWord =>
                new RegistryValueBackup { Name = valueName, Kind = kind, QwordValue = Convert.ToInt64(value) },
            RegistryValueKind.Binary or RegistryValueKind.None =>
                new RegistryValueBackup { Name = valueName, Kind = kind, BinaryValue = value as byte[] ?? [] },
            _ =>
                new RegistryValueBackup { Name = valueName, Kind = kind, StringValue = value.ToString() },
        };
    }

    public void Apply(RegistryKey key)
    {
        object value = Kind switch
        {
            RegistryValueKind.String or RegistryValueKind.ExpandString => StringValue ?? string.Empty,
            RegistryValueKind.MultiString => MultiStringValue ?? [],
            RegistryValueKind.DWord => DwordValue ?? 0,
            RegistryValueKind.QWord => QwordValue ?? 0L,
            RegistryValueKind.Binary or RegistryValueKind.None => BinaryValue ?? [],
            _ => StringValue ?? string.Empty,
        };

        key.SetValue(Name, value, Kind);
    }
}
