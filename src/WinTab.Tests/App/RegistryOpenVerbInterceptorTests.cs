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
                    CommandDefault = "explorer.exe \"%1\""
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
