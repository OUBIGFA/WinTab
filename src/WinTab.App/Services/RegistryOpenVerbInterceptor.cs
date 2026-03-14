using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Win32;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.ShellBridge;

namespace WinTab.App.Services;

public sealed class RegistryOpenVerbInterceptor : IExplorerOpenVerbInterceptor
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
    private const string DelegateExecuteValueName = "DelegateExecute";
    private const string DelegateExecuteComDescription = "WinTab Open Folder DelegateExecute";

    private const string HandlerArgNew = "--wintab-open-folder";
    private const string HandlerArgLegacy = "--open-folder";
    private static readonly string[] TargetClasses =
    [
        "Folder",
        "Directory",
        "Drive"
    ];

    private readonly string _exePath;
    private readonly string _comHostPath;
    private readonly string _comHost32Path;
    private readonly string _comRuntimeConfigPath;
    private readonly string _com32RuntimeConfigPath;
    private readonly Logger _logger;

    public RegistryOpenVerbInterceptor(string exePath, Logger logger)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(exePath);
        _exePath = exePath;
        string baseDirectory = Path.GetDirectoryName(_exePath) ?? string.Empty;
        _comHostPath = Path.Combine(baseDirectory, "WinTab.ShellBridge.comhost.dll");
        _comHost32Path = Path.Combine(baseDirectory, "x86", "WinTab.ShellBridge.comhost.dll");
        _comRuntimeConfigPath = Path.Combine(baseDirectory, "WinTab.ShellBridge.runtimeconfig.json");
        _com32RuntimeConfigPath = Path.Combine(baseDirectory, "x86", "WinTab.ShellBridge.runtimeconfig.json");
        _logger = logger;
    }

    private static string BackupPath => Path.Combine(AppPaths.BaseDirectory, "explorer-open-verb-backup.json");
    private static string DelegateExecuteClsidBraced => "{" + DelegateExecuteIds.OpenFolderDelegateExecuteClsid + "}";

    public static string HandlerArgument => HandlerArgNew;

    public bool IsEnabled()
    {
        try
        {
            bool allVerbsDelegateExecute = true;
            bool allVerbsLegacyCommand = true;

            foreach (string cls in TargetClasses)
            {
                foreach (string verb in TargetVerbs)
                {
                    string? cmd = ReadCommand(cls, verb);
                    string? delegateExecute = ReadDelegateExecute(cls, verb);

                    allVerbsDelegateExecute &= DelegateExecutePointsToWinTab(delegateExecute);
                    allVerbsLegacyCommand &= CommandPointsToWinTab(cmd);
                }
            }

            if (allVerbsDelegateExecute && IsDelegateExecuteComServerRegistered())
                return true;

            return allVerbsLegacyCommand;
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
        RemoveDelegateExecuteComServerRegistration();

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
        catch (Exception ex) when (ex is UnauthorizedAccessException or System.Security.SecurityException or IOException)
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
                DelegateExecute = ReadEffectiveDelegateExecuteValue(cls, verb),
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
        bool comHostExists = File.Exists(_comHostPath);
        bool comHost32Exists = File.Exists(_comHost32Path);
        bool x64RuntimeCompatible = HasCompatibleRuntimeForHost(_comRuntimeConfigPath, RegistryView.Registry64);
        bool x86RuntimeCompatible = HasCompatibleRuntimeForHost(_com32RuntimeConfigPath, RegistryView.Registry32);
        bool preferDelegateExecute = ShouldPreferDelegateExecuteOverride(
            comHostExists,
            comHost32Exists,
            x64RuntimeCompatible,
            x86RuntimeCompatible);
        if (preferDelegateExecute)
        {
            RegisterDelegateExecuteComServer();
        }
        else
        {
            RemoveDelegateExecuteComServerRegistration();
        }

        foreach (string cls in TargetClasses)
        {
            foreach (string verb in TargetVerbs)
            {
                using RegistryKey key = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{cls}\shell\{verb}\command", writable: true)
                    ?? throw new InvalidOperationException("Failed to create registry key.");

                if (preferDelegateExecute)
                {
                    key.SetValue(string.Empty, string.Empty, RegistryValueKind.String);
                    key.SetValue(DelegateExecuteValueName, DelegateExecuteClsidBraced, RegistryValueKind.String);
                }
                else
                {
                    key.DeleteValue(DelegateExecuteValueName, throwOnMissingValue: false);
                    key.SetValue(string.Empty, BuildOverrideCommand(), RegistryValueKind.String);
                }
            }
        }

        // Ensure default verb points to open (standard), but we do not touch other verbs.
        foreach (string cls in TargetClasses)
        {
            using RegistryKey key = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{cls}\shell", writable: true)
                ?? throw new InvalidOperationException("Failed to create registry key.");

            key.SetValue(string.Empty, OpenVerb, RegistryValueKind.String);
        }

        if (preferDelegateExecute)
        {
            _logger.Info("Enabled Explorer open-verb interception via DelegateExecute COM bridge (HKCU, x64+x86).");
        }
        else
        {
            if (comHostExists && !comHost32Exists)
            {
                _logger.Warn($"DelegateExecute x86 COM host is missing, using legacy command override: {_comHost32Path}");
            }
            else if (!comHostExists && comHost32Exists)
            {
                _logger.Warn($"DelegateExecute x64 COM host is missing, using legacy command override: {_comHostPath}");
            }
            else if (!comHostExists && !comHost32Exists)
            {
                _logger.Warn($"DelegateExecute COM hosts are missing, using legacy command override: {_comHostPath} | {_comHost32Path}");
            }
            else if (!x64RuntimeCompatible && !x86RuntimeCompatible)
            {
                _logger.Warn($"DelegateExecute bridge runtimes are incompatible for both architectures, using legacy command override: {_comRuntimeConfigPath} | {_com32RuntimeConfigPath}");
            }
            else if (!x64RuntimeCompatible)
            {
                _logger.Warn($"DelegateExecute x64 runtime is incompatible, using legacy command override: {_comRuntimeConfigPath}");
            }
            else if (!x86RuntimeCompatible)
            {
                _logger.Warn($"DelegateExecute x86 runtime is incompatible, using legacy command override: {_com32RuntimeConfigPath}");
            }
            else
            {
                _logger.Warn("DelegateExecute bridge registration is incomplete, using legacy command override for shell-host compatibility.");
            }

            _logger.Info("Enabled Explorer open-verb interception (legacy command mode, HKCU).");
        }
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

        RemoveDelegateExecuteComServerRegistration();
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
            string? openDelegateExecute = ReadEffectiveDelegateExecuteValue(cls, OpenVerb);
            if (string.IsNullOrWhiteSpace(openCommand) && !DelegateExecutePointsToWinTab(openDelegateExecute))
                return true;

            foreach (string verb in TargetVerbs)
            {
                string? command = ReadEffectiveCommandDefaultValue(cls, verb);
                string? delegateExecute = ReadEffectiveDelegateExecuteValue(cls, verb);

                if (DelegateExecutePointsToWinTab(delegateExecute))
                    continue;

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
                string? delegateExecute = ReadDelegateExecute(cls, verb);
                if (DelegateExecutePointsToWinTab(delegateExecute))
                    return true;

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
                if (entry is null || ShouldDeleteCommandKeyOnRestore(entry.CommandDefault, entry.DelegateExecute))
                {
                    root.DeleteSubKeyTree(commandKey, throwOnMissingSubKey: false);
                }
                else
                {
                    using RegistryKey k = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{commandKey}", writable: true)
                        ?? throw new InvalidOperationException("Failed to create registry key.");

                    if (entry.CommandDefault is null)
                        k.DeleteValue(string.Empty, throwOnMissingValue: false);
                    else
                        k.SetValue(string.Empty, entry.CommandDefault, RegistryValueKind.String);

                    if (entry.DelegateExecute is null)
                        k.DeleteValue(DelegateExecuteValueName, throwOnMissingValue: false);
                    else
                        k.SetValue(DelegateExecuteValueName, entry.DelegateExecute, RegistryValueKind.String);
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

    private string? ReadDelegateExecute(string className, string verb)
    {
        using RegistryKey? k = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{className}\{BuildCommandSubKey(verb)}");
        return k?.GetValue(DelegateExecuteValueName) as string;
    }

    private static string? ReadEffectiveCommandDefaultValue(string className, string verb)
    {
        using RegistryKey? k = Registry.ClassesRoot.OpenSubKey($@"{className}\{BuildCommandSubKey(verb)}", writable: false);
        return k?.GetValue(string.Empty) as string;
    }

    private static string? ReadEffectiveDelegateExecuteValue(string className, string verb)
    {
        using RegistryKey? k = Registry.ClassesRoot.OpenSubKey($@"{className}\{BuildCommandSubKey(verb)}", writable: false);
        return k?.GetValue(DelegateExecuteValueName) as string;
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

    private static bool DelegateExecutePointsToWinTab(string? delegateExecute)
    {
        if (string.IsNullOrWhiteSpace(delegateExecute))
            return false;

        string normalized = delegateExecute.Trim().Trim('{', '}');
        return string.Equals(
            normalized,
            DelegateExecuteIds.OpenFolderDelegateExecuteClsid,
            StringComparison.OrdinalIgnoreCase);
    }

    private string DelegateExecuteClsidKeyPath => $@"Software\Classes\CLSID\{DelegateExecuteClsidBraced}";

    private static RegistryView[] GetDelegateExecuteRegistryViews()
    {
        return Environment.Is64BitOperatingSystem
            ? [RegistryView.Registry64, RegistryView.Registry32]
            : [RegistryView.Registry32];
    }

    private string GetComHostPathForView(RegistryView view)
    {
        return view == RegistryView.Registry32 ? _comHost32Path : _comHostPath;
    }

    private bool IsDelegateExecuteComServerRegistered()
    {
        try
        {
            foreach (RegistryView view in GetDelegateExecuteRegistryViews())
            {
                using RegistryKey? inproc = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view)
                    .OpenSubKey($@"{DelegateExecuteClsidKeyPath}\InProcServer32", writable: false);
                if (inproc is null)
                    return false;

                string? registeredPath = inproc.GetValue(string.Empty) as string;
                if (string.IsNullOrWhiteSpace(registeredPath))
                    return false;

                string expectedPath = GetComHostPathForView(view);
                string normalizedRegistered = Path.GetFullPath(registeredPath).Replace('/', '\\');
                string normalizedExpected = Path.GetFullPath(expectedPath).Replace('/', '\\');

                if (!string.Equals(normalizedRegistered, normalizedExpected, StringComparison.OrdinalIgnoreCase))
                    return false;

                string? threadingModel = inproc.GetValue("ThreadingModel") as string;
                if (!string.Equals(threadingModel, "Apartment", StringComparison.OrdinalIgnoreCase))
                    return false;

                if (!File.Exists(normalizedExpected))
                    return false;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private void RegisterDelegateExecuteComServer()
    {
        foreach (RegistryView view in GetDelegateExecuteRegistryViews())
        {
            string comHostPath = GetComHostPathForView(view);
            using RegistryKey clsidKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view)
                .CreateSubKey(DelegateExecuteClsidKeyPath, writable: true)
                ?? throw new InvalidOperationException("Failed to create DelegateExecute CLSID key.");
            clsidKey.SetValue(string.Empty, DelegateExecuteComDescription, RegistryValueKind.String);

            using RegistryKey inproc = clsidKey.CreateSubKey("InProcServer32", writable: true)
                ?? throw new InvalidOperationException("Failed to create DelegateExecute InProcServer32 key.");
            inproc.SetValue(string.Empty, comHostPath, RegistryValueKind.String);
            inproc.SetValue("ThreadingModel", "Apartment", RegistryValueKind.String);
        }
    }

    private void RemoveDelegateExecuteComServerRegistration()
    {
        foreach (RegistryView view in GetDelegateExecuteRegistryViews())
        {
            using RegistryKey? root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view)
                .OpenSubKey(@"Software\Classes\CLSID", writable: true);
            root?.DeleteSubKeyTree(DelegateExecuteClsidBraced, throwOnMissingSubKey: false);
        }
    }

    private string BuildOverrideCommand()
    {
        // Marker argument distinguishes handler invocation from normal startup.
        return $"\"{_exePath}\" {HandlerArgNew} \"%1\"";
    }

    private static bool HasCompatibleRuntimeForHost(string runtimeConfigPath, RegistryView view)
    {
        if (!File.Exists(runtimeConfigPath))
            return false;

        try
        {
            using JsonDocument document = JsonDocument.Parse(File.ReadAllText(runtimeConfigPath));
            JsonElement runtimeOptions = document.RootElement.GetProperty("runtimeOptions");
            string rollForward = runtimeOptions.TryGetProperty("rollForward", out JsonElement rollForwardElement)
                ? rollForwardElement.GetString() ?? "LatestMinor"
                : "LatestMinor";

            if (!runtimeOptions.TryGetProperty("frameworks", out JsonElement frameworks) || frameworks.ValueKind != JsonValueKind.Array)
                return false;

            foreach (JsonElement framework in frameworks.EnumerateArray())
            {
                string? name = framework.GetProperty("name").GetString();
                string? versionText = framework.GetProperty("version").GetString();
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(versionText))
                    return false;

                if (!Version.TryParse(versionText, out Version? requestedVersion))
                    return false;

                IReadOnlyList<Version> installedVersions = GetInstalledFrameworkVersions(name, view);
                if (!IsFrameworkRuntimeCompatible(requestedVersion, installedVersions, rollForward))
                    return false;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static IReadOnlyList<Version> GetInstalledFrameworkVersions(string frameworkName, RegistryView view)
    {
        string root = view == RegistryView.Registry32
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "dotnet", "shared", frameworkName)
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "shared", frameworkName);

        if (!Directory.Exists(root))
            return [];

        return Directory.GetDirectories(root)
            .Select(Path.GetFileName)
            .Where(static name => !string.IsNullOrWhiteSpace(name))
            .Select(static name => Version.TryParse(name, out Version? version) ? version : null)
            .OfType<Version>()
            .OrderBy(static version => version)
            .ToArray();
    }

    private static bool IsFrameworkRuntimeCompatible(Version requestedVersion, IReadOnlyList<Version> installedVersions, string rollForward)
    {
        return NormalizeRollForward(rollForward) switch
        {
            "major" or "latestmajor" => installedVersions.Any(version =>
                version.Major > requestedVersion.Major ||
                (version.Major == requestedVersion.Major && version >= requestedVersion)),
            _ => installedVersions.Any(version =>
                version.Major == requestedVersion.Major &&
                version.Minor == requestedVersion.Minor &&
                version >= requestedVersion),
        };
    }

    private static string NormalizeRollForward(string? rollForward)
    {
        return string.IsNullOrWhiteSpace(rollForward)
            ? "latestminor"
            : rollForward.Trim().ToLowerInvariant();
    }

    private bool ShouldPreferDelegateExecuteOverride(
        bool comHostExists,
        bool comHost32Exists,
        bool x64RuntimeCompatible,
        bool x86RuntimeCompatible)
    {
        return comHostExists && comHost32Exists && x64RuntimeCompatible && x86RuntimeCompatible;
    }

    private static string ComputeSha256(string text)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        byte[] hash = SHA256.HashData(bytes);
        return Convert.ToHexString(hash);
    }

    private static bool ShouldDeleteCommandKeyOnRestore(string? commandDefault, string? delegateExecute)
    {
        return commandDefault is null && delegateExecute is null;
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
        public string? DelegateExecute { get; init; }
    }
}
