using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Win32;
using WinTab.Diagnostics;
using WinTab.Persistence;

namespace WinTab.App.Services;

public sealed class RegistryOpenVerbInterceptor
{
    private const string OpenVerb = "open";
    private const string NoneVerb = "none";

    private static readonly string[] TargetVerbs =
    [
        "open",
        "explore",
        "opennewwindow"
    ];

    private const string BackupRegistryPath = @"Software\WinTab\Backups\ExplorerOpenVerb";
    private const string BackupRegistryValueName = "BackupJson";

    private const string HandlerArgNew = "--wintab-open-folder";
    private const string HandlerArgLegacy = "--open-folder";
    private static readonly string[] TargetClasses =
    [
        "Folder",
        "Directory",
        "Drive"
    ];

    private readonly string _exePath;
    private readonly Logger _logger;

    public RegistryOpenVerbInterceptor(string exePath, Logger logger)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(exePath);
        _exePath = exePath;
        _logger = logger;
    }

    private static string BackupPath => Path.Combine(AppPaths.BaseDirectory, "explorer-open-verb-backup.json");

    public static string HandlerArgument => HandlerArgNew;

    public bool IsEnabled()
    {
        try
        {
            foreach (string cls in TargetClasses)
            {
                foreach (string verb in TargetVerbs)
                {
                    string? cmd = ReadCommand(cls, verb);
                    if (!CommandPointsToWinTab(cmd))
                        return false;
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public void EnableOrRepair()
    {
        EnsureBackupExists();
        PersistBackupToRegistry();
        WriteOverride();
    }

    public void DisableAndRestore()
    {
        BackupFile? backup = null;

        if (File.Exists(BackupPath))
        {
            try
            {
                backup = LoadBackup();
            }
            catch
            {
                backup = null;
            }
        }

        backup ??= TryLoadBackupFromRegistry();

        if (backup is null)
        {
            _logger.Warn("Explorer open-verb backup is missing; applying safe defaults.");
            ApplySafeDefaultsAndRemoveOverrides();
            return;
        }

        RestoreFromBackup(backup);

        if (IsLikelyBrokenOpenVerbState())
        {
            _logger.Warn("Restored backup produced invalid open-verb state; applying safe defaults.");
            ApplySafeDefaultsAndRemoveOverrides();
        }

        try { File.Delete(BackupPath); } catch { /* ignore */ }
        try { DeleteBackupFromRegistry(); } catch { /* ignore */ }
    }

    public void StartupSelfCheck(bool settingEnabled)
    {
        try
        {
            bool registryPointsToUs = IsEnabled();
            bool registryPointsToAnyWinTabHandler = HasAnyWinTabHandlerOverride();
            bool backupExistsOnDisk = File.Exists(BackupPath);
            bool backupExistsInRegistry = TryLoadBackupFromRegistry() is not null;
            bool hasAnyBackup = backupExistsOnDisk || backupExistsInRegistry;

            if (!settingEnabled && (registryPointsToUs || registryPointsToAnyWinTabHandler))
            {
                _logger.Warn("Explorer open-verb points to WinTab but setting is disabled; restoring defaults.");
                DisableAndRestore();
                return;
            }

            if (!settingEnabled && IsLikelyBrokenOpenVerbState())
            {
                _logger.Warn("Detected invalid Explorer open-verb state while interception is disabled; applying safe defaults.");
                ApplySafeDefaultsAndRemoveOverrides();
                return;
            }

            if (settingEnabled)
            {
                // Unsafe state: override exists but no valid backup source.
                // Reset override first to avoid backing up already-hijacked values.
                if (registryPointsToUs && !hasAnyBackup)
                {
                    _logger.Warn("Explorer open-verb points to WinTab but no backup exists; resetting override before repair.");
                    DisableAndRestore();
                    EnableOrRepair();
                    return;
                }

                // If enabled, we should have both a working override and at least one backup source.
                if (!registryPointsToUs || !hasAnyBackup)
                {
                    _logger.Warn("Explorer open-verb interception is enabled but state is inconsistent; repairing.");
                    EnableOrRepair();
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Error("Explorer open-verb self-check failed.", ex);
        }
    }

    private void EnsureBackupExists()
    {
        if (File.Exists(BackupPath))
            return;

        Directory.CreateDirectory(AppPaths.BaseDirectory);

        BackupFile backup = CreateBackup();
        File.WriteAllText(BackupPath, JsonSerializer.Serialize(backup, new JsonSerializerOptions { WriteIndented = true }));
        _logger.Info($"Saved Explorer open-verb backup: {BackupPath}");
    }

    private BackupFile CreateBackup()
    {
        var backup = new BackupFile
        {
            ExePath = _exePath,
            CreatedAtUtc = DateTimeOffset.UtcNow,
            Entries = TargetClasses.SelectMany(cls => TargetVerbs.Select(verb => new BackupEntry
            {
                ClassName = cls,
                Verb = verb,
                DefaultVerb = ReadEffectiveDefaultVerb(cls),
                CommandDefault = ReadEffectiveCommandDefaultValue(cls, verb),
            })).ToList(),
        };

        backup.Sha256 = ComputeSha256(JsonSerializer.Serialize(backup with { Sha256 = null }));
        return backup;
    }

    private void PersistBackupToRegistry()
    {
        BackupFile backup = File.Exists(BackupPath)
            ? LoadBackup()
            : CreateBackup();

        string json = JsonSerializer.Serialize(backup);

        using RegistryKey key = Registry.CurrentUser.CreateSubKey(BackupRegistryPath, writable: true)
            ?? throw new InvalidOperationException("Failed to create backup registry key.");

        key.SetValue(BackupRegistryValueName, json, RegistryValueKind.String);
    }

    private BackupFile? TryLoadBackupFromRegistry()
    {
        try
        {
            using RegistryKey? key = Registry.CurrentUser.OpenSubKey(BackupRegistryPath, writable: false);
            if (key is null)
                return null;

            string? json = key.GetValue(BackupRegistryValueName) as string;
            if (string.IsNullOrWhiteSpace(json))
                return null;

            var backup = JsonSerializer.Deserialize<BackupFile>(json);
            if (backup is null)
                return null;

            string expected = ComputeSha256(JsonSerializer.Serialize(backup with { Sha256 = null }));
            if (!string.Equals(expected, backup.Sha256, StringComparison.OrdinalIgnoreCase))
                return null;

            return backup;
        }
        catch
        {
            return null;
        }
    }

    private void DeleteBackupFromRegistry()
    {
        using RegistryKey? root = Registry.CurrentUser.OpenSubKey(@"Software", writable: true);
        root?.DeleteSubKeyTree(@"WinTab\Backups\ExplorerOpenVerb", throwOnMissingSubKey: false);
    }

    private BackupFile LoadBackup()
    {
        string json = File.ReadAllText(BackupPath);
        var backup = JsonSerializer.Deserialize<BackupFile>(json) ?? throw new InvalidOperationException("Backup file invalid.");

        string expected = ComputeSha256(JsonSerializer.Serialize(backup with { Sha256 = null }));
        if (!string.Equals(expected, backup.Sha256, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Backup file hash mismatch.");

        return backup;
    }

    private void WriteOverride()
    {
        foreach (string cls in TargetClasses)
        {
            foreach (string verb in TargetVerbs)
            {
                using RegistryKey key = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{cls}\shell\{verb}\command", writable: true)
                    ?? throw new InvalidOperationException("Failed to create registry key.");

                key.SetValue(string.Empty, BuildOverrideCommand(), RegistryValueKind.String);
            }
        }

        // Ensure default verb points to open (standard), but we do not touch other verbs.
        foreach (string cls in TargetClasses)
        {
            using RegistryKey key = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{cls}\shell", writable: true)
                ?? throw new InvalidOperationException("Failed to create registry key.");

            key.SetValue(string.Empty, OpenVerb, RegistryValueKind.String);
        }

        _logger.Info("Enabled Explorer open-verb interception (HKCU)." );
    }

    private void RemoveOverridesOnly()
    {
        using RegistryKey? root = Registry.CurrentUser.OpenSubKey(@"Software\Classes", writable: true);
        if (root is null)
            return;

        foreach (string cls in TargetClasses)
        {
            foreach (string verb in TargetVerbs)
            {
                root.DeleteSubKeyTree($@"{cls}\{BuildCommandSubKey(verb)}", throwOnMissingSubKey: false);
            }
        }
    }

    private void ApplySafeDefaultsAndRemoveOverrides()
    {
        RemoveOverridesOnly();

        foreach (string cls in TargetClasses)
        {
            using RegistryKey? shell = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{cls}\shell", writable: true);
            if (shell is null)
                continue;

            shell.SetValue(string.Empty, GetSafeDefaultVerb(cls), RegistryValueKind.String);
        }

        _logger.Info("Applied safe Explorer open-verb defaults.");
    }

    private static string GetSafeDefaultVerb(string className)
    {
        return className switch
        {
            "Directory" => NoneVerb,
            "Drive" => NoneVerb,
            _ => OpenVerb,
        };
    }

    private bool IsLikelyBrokenOpenVerbState()
    {
        foreach (string cls in TargetClasses)
        {
            string? defaultVerb = ReadEffectiveDefaultVerb(cls);
            if (!string.Equals(defaultVerb, OpenVerb, StringComparison.OrdinalIgnoreCase))
                continue;

            string? openCommand = ReadEffectiveCommandDefaultValue(cls, OpenVerb);
            if (string.IsNullOrWhiteSpace(openCommand))
                return true;

            foreach (string verb in TargetVerbs)
            {
                string? command = ReadCommand(cls, verb);
                if (string.IsNullOrWhiteSpace(command))
                    continue;

                if (!CommandPointsToWinTab(command))
                    return true;
            }
        }

        return false;
    }

    private bool HasAnyWinTabHandlerOverride()
    {
        foreach (string cls in TargetClasses)
        {
            foreach (string verb in TargetVerbs)
            {
                string? command = ReadCommand(cls, verb);
                if (string.IsNullOrWhiteSpace(command))
                    continue;

                if (command.Contains(HandlerArgNew, StringComparison.OrdinalIgnoreCase) ||
                    command.Contains(HandlerArgLegacy, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private void RestoreFromBackup(BackupFile backup)
    {
        using RegistryKey? root = Registry.CurrentUser.OpenSubKey(@"Software\Classes", writable: true);
        if (root is null)
            return;

        var entriesByClass = backup.Entries
            .GroupBy(e => e.ClassName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        foreach (string cls in TargetClasses)
        {
            entriesByClass.TryGetValue(cls, out List<BackupEntry>? classEntries);

            string? defaultVerb = classEntries?.FirstOrDefault()?.DefaultVerb;
            using (RegistryKey? shell = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{cls}\shell", writable: true))
            {
                if (shell is not null)
                {
                    if (defaultVerb is null)
                        shell.DeleteValue(string.Empty, throwOnMissingValue: false);
                    else
                        shell.SetValue(string.Empty, defaultVerb, RegistryValueKind.String);
                }
            }

            foreach (string verb in TargetVerbs)
            {
                BackupEntry? entry = classEntries?.FirstOrDefault(e =>
                    string.Equals(e.Verb, verb, StringComparison.OrdinalIgnoreCase));

                string commandKey = $@"{cls}\shell\{verb}\command";
                if (entry is null || entry.CommandDefault is null)
                {
                    root.DeleteSubKeyTree(commandKey, throwOnMissingSubKey: false);
                }
                else
                {
                    using RegistryKey k = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{commandKey}", writable: true)
                        ?? throw new InvalidOperationException("Failed to create registry key.");

                    k.SetValue(string.Empty, entry.CommandDefault, RegistryValueKind.String);
                }
            }
        }

        _logger.Info("Restored Explorer open-verb defaults from backup." );
    }

    private static string BuildCommandSubKey(string verb) => $@"shell\{verb}\command";

    private string? ReadCommand(string className, string verb)
    {
        using RegistryKey? k = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{className}\{BuildCommandSubKey(verb)}");
        return k?.GetValue(string.Empty) as string;
    }

    private static string? ReadEffectiveCommandDefaultValue(string className, string verb)
    {
        using RegistryKey? k = Registry.ClassesRoot.OpenSubKey($@"{className}\{BuildCommandSubKey(verb)}", writable: false);
        return k?.GetValue(string.Empty) as string;
    }

    private static string? ReadEffectiveDefaultVerb(string className)
    {
        using RegistryKey? k = Registry.ClassesRoot.OpenSubKey($@"{className}\shell", writable: false);
        return k?.GetValue(string.Empty) as string;
    }

    private bool CommandPointsToWinTab(string? command)
    {
        if (string.IsNullOrWhiteSpace(command))
            return false;

        if (!command.Contains(HandlerArgNew, StringComparison.OrdinalIgnoreCase) &&
            !command.Contains(HandlerArgLegacy, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        string normalized = command.Trim();
        string quotedExe = $"\"{_exePath}\"";

        if (normalized.Contains(quotedExe, StringComparison.OrdinalIgnoreCase))
            return true;

        // Also allow unquoted exact path formats for compatibility.
        if (normalized.StartsWith(_exePath + " ", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalized, _exePath, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private string BuildOverrideCommand()
    {
        // Marker argument distinguishes handler invocation from normal startup.
        return $"\"{_exePath}\" {HandlerArgNew} \"%1\"";
    }

    private static string ComputeSha256(string text)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        byte[] hash = SHA256.HashData(bytes);
        return Convert.ToHexString(hash);
    }

    private sealed record BackupFile
    {
        public required string ExePath { get; init; }
        public required DateTimeOffset CreatedAtUtc { get; init; }
        public required List<BackupEntry> Entries { get; init; }
        public string? Sha256 { get; set; }
    }

    private sealed record BackupEntry
    {
        public required string ClassName { get; init; }
        public required string Verb { get; init; }
        public string? DefaultVerb { get; init; }
        public string? CommandDefault { get; init; }
    }
}
