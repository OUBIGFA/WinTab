using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using FluentAssertions;
using WinTab.App.Services;
using WinTab.Diagnostics;
using Xunit;

namespace WinTab.Tests.App;

public sealed class RegistryOpenVerbInterceptorTests : IDisposable
{
    private readonly string _tempDir;
    private readonly Logger _logger;
    private readonly string _backupPath;
    private readonly string? _originalBackupContent;
    private readonly bool _hadOriginalBackup;

    public RegistryOpenVerbInterceptorTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "WinTabRegistryInterceptorTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _logger = new Logger(Path.Combine(_tempDir, "test.log"));

        _backupPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WinTab", "explorer-open-verb-backup.json");
        _hadOriginalBackup = File.Exists(_backupPath);
        _originalBackupContent = _hadOriginalBackup ? File.ReadAllText(_backupPath) : null;
    }

    [Fact]
    public void CommandPointsToWinTab_ShouldAcceptQuotedExeAndNewHandlerArg()
    {
        string exePath = @"C:\Program Files\WinTab\WinTab.exe";
        var interceptor = new RegistryOpenVerbInterceptor(exePath, _logger);

        string command = @"""C:\Program Files\WinTab\WinTab.exe"" --wintab-open-folder ""%1""";
        bool matches = InvokeCommandPointsToWinTab(interceptor, command);

        matches.Should().BeTrue();
    }

    [Fact]
    public void CommandPointsToWinTab_ShouldAcceptUnquotedExeWithLegacyArg()
    {
        string exePath = @"C:\WinTab\WinTab.exe";
        var interceptor = new RegistryOpenVerbInterceptor(exePath, _logger);

        string command = @"C:\WinTab\WinTab.exe --open-folder ""%1""";
        bool matches = InvokeCommandPointsToWinTab(interceptor, command);

        matches.Should().BeTrue();
    }

    [Fact]
    public void CommandPointsToWinTab_ShouldRejectCommandWithoutHandlerArg()
    {
        string exePath = @"C:\WinTab\WinTab.exe";
        var interceptor = new RegistryOpenVerbInterceptor(exePath, _logger);

        string command = @"""C:\WinTab\WinTab.exe"" ""%1""";
        bool matches = InvokeCommandPointsToWinTab(interceptor, command);

        matches.Should().BeFalse();
    }

    [Fact]
    public void CommandPointsToWinTab_ShouldRejectOtherExecutable()
    {
        string exePath = @"C:\WinTab\WinTab.exe";
        var interceptor = new RegistryOpenVerbInterceptor(exePath, _logger);

        string command = @"""C:\Other\Other.exe"" --wintab-open-folder ""%1""";
        bool matches = InvokeCommandPointsToWinTab(interceptor, command);

        matches.Should().BeFalse();
    }

    [Fact]
    public void LoadBackup_ShouldRejectHashMismatch()
    {
        string exePath = @"C:\WinTab\WinTab.exe";
        var interceptor = new RegistryOpenVerbInterceptor(exePath, _logger);

        string backupPath = GetBackupPath();
        Directory.CreateDirectory(Path.GetDirectoryName(backupPath)!);
        File.WriteAllText(backupPath, BuildBackupJsonWithHash(exePath, "BAD_HASH"));

        Action act = () => InvokeLoadBackup(interceptor);

        act.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidOperationException>()
            .WithMessage("*hash mismatch*");
    }

    [Fact]
    public void LoadBackup_ShouldAcceptValidHash()
    {
        string exePath = @"C:\WinTab\WinTab.exe";
        var interceptor = new RegistryOpenVerbInterceptor(exePath, _logger);

        string backupPath = GetBackupPath();
        Directory.CreateDirectory(Path.GetDirectoryName(backupPath)!);

        DateTimeOffset createdAt = DateTimeOffset.UtcNow;
        string jsonWithoutHash = BuildBackupJsonWithHash(exePath, null, createdAt);
        string validHash = ComputeSha256(jsonWithoutHash);
        File.WriteAllText(backupPath, BuildBackupJsonWithHash(exePath, validHash, createdAt));

        object? backup = InvokeLoadBackup(interceptor);
        backup.Should().NotBeNull();
    }

    [Fact]
    public void RestorePolicy_WhenOnlyDelegateExecuteExists_ShouldKeepCommandKey()
    {
        MethodInfo? method = typeof(RegistryOpenVerbInterceptor).GetMethod(
            "ShouldDeleteCommandKeyOnRestore",
            BindingFlags.NonPublic | BindingFlags.Static);

        method.Should().NotBeNull("restore policy must preserve DelegateExecute-only handlers");

        object? shouldDelete = method?.Invoke(null, [null, "{A1111111-2222-3333-4444-555555555555}"]);
        shouldDelete.Should().BeOfType<bool>();
        ((bool)shouldDelete!).Should().BeFalse(
            "when DelegateExecute is present, restore must not drop the shell handler");
    }

    [Fact]
    public void RuntimeInterceptorSource_ShouldRegisterDelegateExecuteInLocalMachineAndCurrentUserHives()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "RegistryOpenVerbInterceptor.cs"]));

        source.Should().Contain("RegistryHive.LocalMachine",
            "runtime repair must create machine-wide COM registration so Windows 11 Start Menu and elevated shell hosts do not fail with class-not-registered");
        source.Should().Contain("RegistryHive.CurrentUser",
            "runtime repair should still preserve user-scope registration for same-user Explorer scenarios");
    }

    [Fact]
    public void RuntimeInterceptorSource_ShouldRemoveDelegateExecuteFromLocalMachineAndCurrentUserHives()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "RegistryOpenVerbInterceptor.cs"]));

        source.Should().Contain("Failed to remove DelegateExecute COM registration from HKLM",
            "cleanup logging should explicitly cover machine-wide registration removal failures");
        source.Should().Contain("rootLm?.DeleteSubKeyTree",
            "runtime cleanup must attempt to delete the machine-wide COM registration");
        source.Should().Contain("rootCu?.DeleteSubKeyTree",
            "runtime cleanup must still delete the user-scope COM registration");
    }

    [Fact]
    public void RuntimeInterceptorSource_ShouldUseVolatileSessionOnlyOverridesWhenPersistenceAcrossRebootIsDisabled()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "RegistryOpenVerbInterceptor.cs"]));
        string helperSource = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.Platform.Win32", "VolatileRegistryKeyFactory.cs"]));

        source.Should().Contain("WriteSessionOnlyOverride",
            "session-only startup should not rely on persistent HKCU Classes overrides");
        source.Should().Contain("persistAcrossReboot",
            "the interceptor must branch between reboot-persistent and session-only shell registration");
        helperSource.Should().Contain("REG_OPTION_VOLATILE",
            "session-only shell overrides must be created as volatile registry keys so Windows reboot clears them automatically");
    }

    [Fact]
    public void RuntimeInterceptorSource_ShouldPreferVolatileDelegateExecuteForSessionOnlyMode()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "RegistryOpenVerbInterceptor.cs"]));

        source.Should().Contain("RegisterDelegateExecuteComServerVolatileCurrentUser",
            "session-only interception should use the direct DelegateExecute bridge first so folder opens can be reused without flashing a temporary Explorer window");
        source.Should().Contain("volatile DelegateExecute bridge",
            "logging should make it explicit when the no-flicker session-only DelegateExecute path is active");
        source.Should().Contain("volatile HKCU command overrides",
            "legacy command mode should remain only as the compatibility fallback when the DelegateExecute bridge is unavailable");
    }

    public void Dispose()
    {
        _logger.Dispose();

        try
        {
            if (_hadOriginalBackup)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_backupPath)!);
                File.WriteAllText(_backupPath, _originalBackupContent!);
            }
            else if (File.Exists(_backupPath))
            {
                File.Delete(_backupPath);
            }
        }
        catch
        {
            // ignore
        }

        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // ignore
        }
    }

    private static bool InvokeCommandPointsToWinTab(RegistryOpenVerbInterceptor interceptor, string? command)
    {
        MethodInfo method = typeof(RegistryOpenVerbInterceptor).GetMethod("CommandPointsToWinTab", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Method not found: CommandPointsToWinTab");

        object? result = method.Invoke(interceptor, [command]);
        return result is bool b && b;
    }

    private static object? InvokeLoadBackup(RegistryOpenVerbInterceptor interceptor)
    {
        MethodInfo method = typeof(RegistryOpenVerbInterceptor).GetMethod("LoadBackup", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Method not found: LoadBackup");

        return method.Invoke(interceptor, null);
    }

    private static string GetBackupPath()
    {
        PropertyInfo property = typeof(RegistryOpenVerbInterceptor).GetProperty("BackupPath", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Property not found: BackupPath");

        return property.GetValue(null) as string
            ?? throw new InvalidOperationException("BackupPath is null.");
    }

    private static string BuildBackupJsonWithHash(string exePath, string? hash, DateTimeOffset? createdAtUtc = null)
    {
        var payload = new
        {
            ExePath = exePath,
            CreatedAtUtc = createdAtUtc ?? DateTimeOffset.UtcNow,
            Entries = new[]
            {
                new
                {
                    ClassName = "Folder",
                    Verb = "open",
                    DefaultVerb = "open",
                    CommandDefault = "explorer.exe \"%1\"",
                    DelegateExecute = (string?)null
                }
            },
            Sha256 = hash
        };

        return JsonSerializer.Serialize(payload);
    }

    private static string ComputeSha256(string text)
    {
        MethodInfo method = typeof(RegistryOpenVerbInterceptor).GetMethod("ComputeSha256", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Method not found: ComputeSha256");

        object? hash = method.Invoke(null, [text]);
        return hash as string ?? throw new InvalidOperationException("ComputeSha256 returned null.");
    }
}
