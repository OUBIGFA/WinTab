using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;
using WinTab.ShellBridge;

namespace WinTab.App.Services;

public sealed class RegistryOpenVerbInterceptor : IExplorerOpenVerbInterceptor
{
    private const string OpenVerb = "open";
    private const string DelegateExecuteValueName = "DelegateExecute";
    private const string DelegateExecuteComDescription = "WinTab Open Folder DelegateExecute";
    private const string HandlerArgNew = "--wintab-open-folder";
    private const string HandlerArgLegacy = "--open-folder";
    private const string CompanionArg = "--wintab-companion";
    private const string CompanionWatchParentArg = "--watch-parent";
    private const uint ShcneAssocChanged = 0x08000000;
    private const uint ShcnfIdList = 0x0000;

    private static readonly string[] TargetVerbs =
    [
        "open",
        "explore",
        "opennewwindow"
    ];

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
    private readonly ExplorerShellBaselineStore _baselineStore;

    public RegistryOpenVerbInterceptor(string exePath, Logger logger)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(exePath);
        ArgumentNullException.ThrowIfNull(logger);

        _exePath = exePath;
        string baseDirectory = Path.GetDirectoryName(_exePath) ?? string.Empty;
        _comHostPath = Path.Combine(baseDirectory, "WinTab.ShellBridge.comhost.dll");
        _comHost32Path = Path.Combine(baseDirectory, "x86", "WinTab.ShellBridge.comhost.dll");
        _comRuntimeConfigPath = Path.Combine(baseDirectory, "WinTab.ShellBridge.runtimeconfig.json");
        _com32RuntimeConfigPath = Path.Combine(baseDirectory, "x86", "WinTab.ShellBridge.runtimeconfig.json");
        _logger = logger;
        _baselineStore = new ExplorerShellBaselineStore(exePath, logger);
    }

    private static string BackupPath => ExplorerShellBaselineStore.BackupPath;
    private static string DelegateExecuteClsidBraced => "{" + DelegateExecuteIds.OpenFolderDelegateExecuteClsid + "}";
    private static string MalformedDelegateExecuteClsidBraced => DelegateExecuteClsidBraced + "}";

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
                    string? command = ReadCommand(cls, verb);
                    string? delegateExecute = ReadDelegateExecute(cls, verb);
                    allVerbsDelegateExecute &= DelegateExecutePointsToWinTab(delegateExecute);
                    allVerbsLegacyCommand &= CommandPointsToWinTab(command);
                }
            }

            if (allVerbsDelegateExecute && IsDelegateExecuteComServerRegistered(requireMachineWideRegistration: true))
                return true;

            return allVerbsLegacyCommand;
        }
        catch
        {
            return false;
        }
    }

    public void EnableOrRepair(bool persistAcrossReboot)
    {
        _baselineStore.EnsureBaselineExists();
        WriteOverride(persistAcrossReboot);
    }

    public void DisableAndRestore(bool deleteBackup = true)
    {
        ExplorerShellBaseline? baseline = _baselineStore.TryLoadBaseline();
        if (baseline is null)
        {
            _logger.Warn("Explorer shell baseline is missing; removing WinTab overrides only.");
            ApplySafeDefaultsAndRemoveOverrides();
            if (deleteBackup)
                _baselineStore.DeletePersistedBaseline();
            return;
        }

        RestoreFromBackup(baseline);
        RemoveCurrentUserDelegateExecuteComServerRegistration();

        if (IsLikelyBrokenOpenVerbState())
        {
            _logger.Warn("Restored Explorer shell baseline still looks inconsistent; removing WinTab overrides only.");
            ApplySafeDefaultsAndRemoveOverrides();
        }
        else
        {
            NotifyShellAssociationChanged();
        }

        if (deleteBackup)
            _baselineStore.DeletePersistedBaseline();
    }

    public void StartupSelfCheck(bool settingEnabled, bool persistAcrossReboot)
    {
        try
        {
            bool registryPointsToUs = IsEnabled();
            bool registryPointsToAnyWinTabHandler = HasAnyWinTabHandlerOverride();
            bool hasAnyBackup = _baselineStore.TryLoadBaseline() is not null;

            if (!settingEnabled && (registryPointsToUs || registryPointsToAnyWinTabHandler))
            {
                _logger.Warn("Explorer open-verb points to WinTab while the feature is disabled; restoring baseline.");
                DisableAndRestore(deleteBackup: false);
                return;
            }

            if (!settingEnabled && IsLikelyBrokenOpenVerbState())
            {
                _logger.Warn("Detected invalid Explorer open-verb state while interception is disabled; removing WinTab overrides.");
                ApplySafeDefaultsAndRemoveOverrides();
                return;
            }

            if (!settingEnabled)
                return;

            if (persistAcrossReboot)
                _logger.Warn("Persistent Explorer interception is no longer supported; runtime-only mode will be used.");

            if (ShouldResetOverrideBeforeRepair(registryPointsToUs, registryPointsToAnyWinTabHandler, hasAnyBackup))
            {
                _logger.Warn("Detected residual WinTab shell override without a trusted baseline; clearing override before re-enabling.");
                DisableAndRestore(deleteBackup: false);
                return;
            }

            if (!registryPointsToUs || !hasAnyBackup)
            {
                _logger.Warn("Explorer open-verb interception is enabled but state is inconsistent; repairing in runtime-only mode.");
                EnableOrRepair(persistAcrossReboot: false);
            }
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or System.Security.SecurityException or IOException)
        {
            _logger.Error("Explorer open-verb self-check failed.", ex);
        }
    }

    private void WriteOverride(bool persistAcrossReboot)
    {
        if (persistAcrossReboot)
            _logger.Warn("Persistent Explorer interception was requested, but WinTab now enforces runtime-only interception for safety.");

        WriteSessionOnlyOverride();
    }

    private void WriteSessionOnlyOverride()
    {
        RemoveOverridesOnly();

        if (!EnsureCompanionCleanupReady())
        {
            _logger.Warn("Companion cleanup process could not be started; Explorer interception will remain disabled.");
            return;
        }

        bool comHostExists = File.Exists(_comHostPath);
        bool comHost32Exists = File.Exists(_comHost32Path);
        bool x64RuntimeCompatible = HasCompatibleRuntimeForHost(_comRuntimeConfigPath, RegistryView.Registry64);
        bool x86RuntimeCompatible = HasCompatibleRuntimeForHost(_com32RuntimeConfigPath, RegistryView.Registry32);
        bool delegateExecutePrerequisitesSatisfied = ShouldPreferDelegateExecuteOverride(
            comHostExists,
            comHost32Exists,
            x64RuntimeCompatible,
            x86RuntimeCompatible);

        bool machineWideRegistrationReady = delegateExecutePrerequisitesSatisfied &&
            IsDelegateExecuteComServerRegistered(requireMachineWideRegistration: true);

        if (delegateExecutePrerequisitesSatisfied && !machineWideRegistrationReady)
        {
            _logger.Warn("DelegateExecute machine-wide COM bridge is unavailable, using session-only legacy command override for shell-host compatibility.");
        }

        bool preferDelegateExecute = ShouldUseDelegateExecuteOverride(
            delegateExecutePrerequisitesSatisfied,
            machineWideRegistrationReady);

        if (preferDelegateExecute)
        {
            WriteSessionOnlyVerbOverrides(preferDelegateExecute: true);
            NotifyShellAssociationChanged();
            _logger.Info("Enabled Explorer open-verb interception in runtime-only mode via machine-wide DelegateExecute bridge.");
            return;
        }

        WriteSessionOnlyVerbOverrides(preferDelegateExecute: false);
        LogLegacyDelegateExecuteFallbackReason(comHostExists, comHost32Exists, x64RuntimeCompatible, x86RuntimeCompatible);
        NotifyShellAssociationChanged();
        _logger.Info("Enabled Explorer open-verb interception in runtime-only mode via volatile HKCU command overrides.");
    }

    private bool EnsureCompanionCleanupReady()
    {
        try
        {
            if (!File.Exists(_exePath))
            {
                _logger.Warn($"Companion cleanup launcher path is missing: {_exePath}");
                return false;
            }

            Process? process = Process.Start(new ProcessStartInfo
            {
                FileName = _exePath,
                Arguments = $"{CompanionArg} {CompanionWatchParentArg} {Environment.ProcessId}",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
            });

            return process is not null;
        }
        catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception or FileNotFoundException)
        {
            _logger.Warn($"Failed to start companion cleanup process: {ex.Message}");
            return false;
        }
    }

    private void WriteSessionOnlyVerbOverrides(bool preferDelegateExecute)
    {
        foreach (string cls in TargetClasses)
        {
            foreach (RegistryView view in GetDelegateExecuteRegistryViews())
            {
                using RegistryKey shellKey = VolatileRegistryKeyFactory.CreateCurrentUserVolatileSubKey(
                    $@"Software\Classes\{cls}\shell",
                    view);
                shellKey.SetValue(string.Empty, OpenVerb, RegistryValueKind.String);

                foreach (string verb in TargetVerbs)
                {
                    using RegistryKey key = VolatileRegistryKeyFactory.CreateCurrentUserVolatileSubKey(
                        $@"Software\Classes\{cls}\shell\{verb}\command",
                        view);

                    if (preferDelegateExecute)
                    {
                        key.SetValue(string.Empty, string.Empty, RegistryValueKind.String);
                        key.SetValue(DelegateExecuteValueName, DelegateExecuteClsidBraced, RegistryValueKind.String);
                    }
                    else
                    {
                        key.SetValue(string.Empty, BuildOverrideCommand(), RegistryValueKind.String);
                        key.DeleteValue(DelegateExecuteValueName, throwOnMissingValue: false);
                    }
                }
            }
        }
    }

    private void LogLegacyDelegateExecuteFallbackReason(
        bool comHostExists,
        bool comHost32Exists,
        bool x64RuntimeCompatible,
        bool x86RuntimeCompatible)
    {
        if (comHostExists && !comHost32Exists)
        {
            _logger.Warn($"DelegateExecute x86 COM host is missing, using session-only legacy command override: {_comHost32Path}");
        }
        else if (!comHostExists && comHost32Exists)
        {
            _logger.Warn($"DelegateExecute x64 COM host is missing, using session-only legacy command override: {_comHostPath}");
        }
        else if (!comHostExists && !comHost32Exists)
        {
            _logger.Warn($"DelegateExecute COM hosts are missing, using session-only legacy command override: {_comHostPath} | {_comHost32Path}");
        }
        else if (!x64RuntimeCompatible && !x86RuntimeCompatible)
        {
            _logger.Warn($"DelegateExecute bridge runtimes are incompatible for both architectures, using session-only legacy command override: {_comRuntimeConfigPath} | {_com32RuntimeConfigPath}");
        }
        else if (!x64RuntimeCompatible)
        {
            _logger.Warn($"DelegateExecute x64 runtime is incompatible, using session-only legacy command override: {_comRuntimeConfigPath}");
        }
        else if (!x86RuntimeCompatible)
        {
            _logger.Warn($"DelegateExecute x86 runtime is incompatible, using session-only legacy command override: {_com32RuntimeConfigPath}");
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
                root.DeleteSubKeyTree($@"{cls}\{BuildCommandSubKey(verb)}", throwOnMissingSubKey: false);
        }

        RemoveCurrentUserDelegateExecuteComServerRegistration();
    }

    private void ApplySafeDefaultsAndRemoveOverrides()
    {
        RemoveOverridesOnly();

        using RegistryKey? root = Registry.CurrentUser.OpenSubKey(@"Software\Classes", writable: true);
        if (root is null)
            return;

        foreach (string cls in TargetClasses)
        {
            using RegistryKey? shell = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{cls}\shell", writable: true);
            shell?.DeleteValue(string.Empty, throwOnMissingValue: false);

            foreach (string verb in TargetVerbs)
                root.DeleteSubKeyTree($@"{cls}\shell\{verb}", throwOnMissingSubKey: false);

            TryDeleteEmptyKey(root, $@"{cls}\shell");
            TryDeleteEmptyKey(root, cls);
        }

        NotifyShellAssociationChanged();
        _logger.Info("Removed WinTab overrides and revealed the native Explorer handlers.");
    }

    private static void TryDeleteEmptyKey(RegistryKey root, string subKeyPath)
    {
        using RegistryKey? key = root.OpenSubKey(subKeyPath, writable: true);
        if (key is null)
            return;

        if (key.SubKeyCount == 0 && key.ValueCount == 0)
            root.DeleteSubKeyTree(subKeyPath, throwOnMissingSubKey: false);
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

    private void RestoreFromBackup(ExplorerShellBaseline baseline)
    {
        _baselineStore.RestoreCurrentUserBaseline(baseline);
        _logger.Info("Restored Explorer shell baseline from backup.");
    }

    private static string BuildCommandSubKey(string verb) => $@"shell\{verb}\command";

    private string? ReadCommand(string className, string verb)
    {
        using RegistryKey? key = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{className}\{BuildCommandSubKey(verb)}");
        return key?.GetValue(string.Empty) as string;
    }

    private string? ReadDelegateExecute(string className, string verb)
    {
        using RegistryKey? key = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{className}\{BuildCommandSubKey(verb)}");
        return key?.GetValue(DelegateExecuteValueName) as string;
    }

    private static string? ReadEffectiveCommandDefaultValue(string className, string verb)
    {
        return ReadEffectiveStringValue($@"Software\Classes\{className}\{BuildCommandSubKey(verb)}", string.Empty);
    }

    private static string? ReadEffectiveDelegateExecuteValue(string className, string verb)
    {
        return ReadEffectiveStringValue($@"Software\Classes\{className}\{BuildCommandSubKey(verb)}", DelegateExecuteValueName);
    }

    private static string? ReadEffectiveDefaultVerb(string className)
    {
        return ReadEffectiveStringValue($@"Software\Classes\{className}\shell", string.Empty);
    }

    private static string? ReadEffectiveStringValue(string subKeyPath, string valueName)
    {
        return ReadStringValue(RegistryHive.CurrentUser, RegistryView.Default, subKeyPath, valueName)
            ?? ReadStringValue(RegistryHive.LocalMachine, RegistryView.Default, subKeyPath, valueName);
    }

    private static string? ReadStringValue(RegistryHive hive, RegistryView view, string subKeyPath, string valueName)
    {
        using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, view);
        using RegistryKey? key = baseKey.OpenSubKey(subKeyPath, writable: false);
        return key?.GetValue(valueName) as string;
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
        return normalized.Contains(quotedExe, StringComparison.OrdinalIgnoreCase) ||
               normalized.StartsWith(_exePath + " ", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(normalized, _exePath, StringComparison.OrdinalIgnoreCase);
    }

    private static bool DelegateExecutePointsToWinTab(string? delegateExecute)
    {
        if (string.IsNullOrWhiteSpace(delegateExecute))
            return false;

        string normalized = delegateExecute.Trim().Trim('{', '}');
        return string.Equals(normalized, DelegateExecuteIds.OpenFolderDelegateExecuteClsid, StringComparison.OrdinalIgnoreCase);
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

    private bool IsDelegateExecuteComServerRegistered(bool requireMachineWideRegistration)
    {
        bool machineWideRegistered = IsDelegateExecuteComServerRegistered(RegistryHive.LocalMachine);
        if (requireMachineWideRegistration)
            return machineWideRegistered;

        return machineWideRegistered || IsDelegateExecuteComServerRegistered(RegistryHive.CurrentUser);
    }

    private bool IsDelegateExecuteComServerRegistered(RegistryHive hive)
    {
        try
        {
            foreach (RegistryView view in GetDelegateExecuteRegistryViews())
            {
                using RegistryKey? inproc = RegistryKey.OpenBaseKey(hive, view)
                    .OpenSubKey($@"{DelegateExecuteClsidKeyPath}\InProcServer32", writable: false);
                if (inproc is null)
                    return false;

                string? registeredPath = inproc.GetValue(string.Empty) as string;
                string expectedPath = GetComHostPathForView(view);
                if (!string.Equals(
                        Path.GetFullPath(registeredPath ?? string.Empty).Replace('/', '\\'),
                        Path.GetFullPath(expectedPath).Replace('/', '\\'),
                        StringComparison.OrdinalIgnoreCase))
                {
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

    private void RemoveCurrentUserDelegateExecuteComServerRegistration()
    {
        foreach (RegistryView view in GetDelegateExecuteRegistryViews())
        {
            using RegistryKey? rootCu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view)
                .OpenSubKey(@"Software\Classes\CLSID", writable: true);
            rootCu?.DeleteSubKeyTree(DelegateExecuteClsidBraced, throwOnMissingSubKey: false);
            rootCu?.DeleteSubKeyTree(MalformedDelegateExecuteClsidBraced, throwOnMissingSubKey: false);
        }
    }

    private void RemoveLegacyDelegateExecuteComServerRegistration()
    {
        foreach (RegistryView view in GetDelegateExecuteRegistryViews())
        {
            try
            {
                using RegistryKey? rootLm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view)
                    .OpenSubKey(@"Software\Classes\CLSID", writable: true);
                rootLm?.DeleteSubKeyTree(DelegateExecuteClsidBraced, throwOnMissingSubKey: false);
                rootLm?.DeleteSubKeyTree(MalformedDelegateExecuteClsidBraced, throwOnMissingSubKey: false);
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException or System.Security.SecurityException)
            {
                _logger.Warn($"Failed to remove legacy DelegateExecute COM registration from HKLM (view {view}): {ex.Message}");
            }

            using RegistryKey? rootCu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view)
                .OpenSubKey(@"Software\Classes\CLSID", writable: true);
            rootCu?.DeleteSubKeyTree(DelegateExecuteClsidBraced, throwOnMissingSubKey: false);
            rootCu?.DeleteSubKeyTree(MalformedDelegateExecuteClsidBraced, throwOnMissingSubKey: false);
        }
    }

    private string BuildOverrideCommand()
    {
        return $"\"{_exePath}\" {HandlerArgNew} \"%1\"";
    }

    private static bool HasCompatibleRuntimeForHost(string runtimeConfigPath, RegistryView view)
    {
        if (!File.Exists(runtimeConfigPath))
            return false;

        try
        {
            using System.Text.Json.JsonDocument document = System.Text.Json.JsonDocument.Parse(File.ReadAllText(runtimeConfigPath));
            System.Text.Json.JsonElement runtimeOptions = document.RootElement.GetProperty("runtimeOptions");
            string rollForward = runtimeOptions.TryGetProperty("rollForward", out System.Text.Json.JsonElement rollForwardElement)
                ? rollForwardElement.GetString() ?? "LatestMinor"
                : "LatestMinor";

            if (!runtimeOptions.TryGetProperty("frameworks", out System.Text.Json.JsonElement frameworks) || frameworks.ValueKind != System.Text.Json.JsonValueKind.Array)
                return false;

            foreach (System.Text.Json.JsonElement framework in frameworks.EnumerateArray())
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

    private static bool ShouldUseDelegateExecuteOverride(
        bool delegateExecutePrerequisitesSatisfied,
        bool machineWideRegistrationReady)
    {
        return delegateExecutePrerequisitesSatisfied && machineWideRegistrationReady;
    }

    private static bool ShouldResetOverrideBeforeRepair(
        bool registryPointsToUs,
        bool registryPointsToAnyWinTabHandler,
        bool hasAnyBackup)
    {
        _ = registryPointsToUs;
        return registryPointsToAnyWinTabHandler && !hasAnyBackup;
    }

    private void NotifyShellAssociationChanged()
    {
        try
        {
            NativeMethods.SHChangeNotify(ShcneAssocChanged, ShcnfIdList, IntPtr.Zero, IntPtr.Zero);
        }
        catch (Exception ex)
        {
            _logger.Warn($"Failed to notify shell association change: {ex.Message}");
        }
    }
}
