using System;
using System.IO;
using System.Reflection;
using FluentAssertions;
using WinTab.App.Services;
using WinTab.Diagnostics;
using Xunit;

namespace WinTab.Tests.App;

public sealed class RegistryOpenVerbInterceptorTests : IDisposable
{
    private readonly string _tempDir;
    private readonly Logger _logger;

    public RegistryOpenVerbInterceptorTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "WinTabRegistryInterceptorTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _logger = new Logger(Path.Combine(_tempDir, "test.log"));
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
    public void RuntimeInterceptorSource_ShouldRemoveDelegateExecuteFromLocalMachineAndCurrentUserHives()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "RegistryOpenVerbInterceptor.cs"]));

        source.Should().Contain("Failed to remove legacy DelegateExecute COM registration from HKLM",
            "cleanup logging should explicitly cover machine-wide legacy registration removal failures");
        source.Should().Contain("rootLm?.DeleteSubKeyTree",
            "runtime cleanup must attempt to delete the machine-wide legacy COM registration");
        source.Should().Contain("rootCu?.DeleteSubKeyTree",
            "runtime cleanup must still delete the user-scope registration");
    }

    [Fact]
    public void RuntimeInterceptorSource_ShouldUseVolatileSessionOnlyOverridesWhenPersistenceAcrossRebootIsDisabled()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "RegistryOpenVerbInterceptor.cs"]));
        string helperSource = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.Platform.Win32", "VolatileRegistryKeyFactory.cs"]));

        source.Should().Contain("WriteSessionOnlyOverride",
            "session-only startup should not rely on persistent HKCU Classes overrides");
        source.Should().Contain("persistAcrossReboot",
            "the interceptor must explicitly reject reboot-persistent shell registration");
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

    [Fact]
    public void RuntimeInterceptorSource_ShouldDeleteUserOverridesWhenBackupIsMissing()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "RegistryOpenVerbInterceptor.cs"]));

        source.Should().Contain("shell?.DeleteValue(string.Empty, throwOnMissingValue: false);",
            "safe fallback should remove the user-scope default verb override when no backup is available");
        source.Should().Contain("root.DeleteSubKeyTree($@\"{cls}\\shell\\{verb}\", throwOnMissingSubKey: false);",
            "safe fallback should delete the user-scope open/explore/opennewwindow overrides so Explorer falls back to the native defaults");
        source.Should().Contain("TryDeleteEmptyKey(root, $@\"{cls}\\shell\");",
            "safe fallback should clean up empty HKCU shell keys after removing WinTab overrides");
        source.Should().Contain("if (!string.Equals(defaultVerb, OpenVerb, StringComparison.OrdinalIgnoreCase))",
            "broken-state detection should keep the same gating semantics as the reference version");
    }

    [Fact]
    public void RuntimeInterceptorSource_ShouldNotUseMergedClassesRootForBaselineCapture()
    {
        string interceptorSource = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "RegistryOpenVerbInterceptor.cs"]));
        string baselineStoreSource = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "ExplorerShellBaselineStore.cs"]));

        interceptorSource.Should().NotContain("Registry.ClassesRoot.OpenSubKey",
            "runtime interception must not read the merged HKCR view when deciding what the native shell state is");
        baselineStoreSource.Should().NotContain("Registry.ClassesRoot.OpenSubKey",
            "baseline capture must read explicit HKCU/HKLM source hives instead of the merged HKCR view, otherwise WinTab can back up its own hijacked state as the 'default'");
    }

    [Fact]
    public void RuntimeInterceptorSource_ShouldUseBaselineStoreAndCompanionCleanup()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "Services", "RegistryOpenVerbInterceptor.cs"]));

        source.Should().Contain("ExplorerShellBaselineStore",
            "runtime interception must snapshot and restore the machine-specific Explorer baseline instead of hard-coding default verbs");
        source.Should().Contain("EnsureCompanionCleanupReady",
            "runtime interception must refuse to hijack Explorer unless a companion cleanup process is in place");
        source.Should().Contain("Companion cleanup process could not be started; Explorer interception will remain disabled.",
            "safety-first startup must fail closed if crash cleanup cannot be armed");
    }

    public void Dispose()
    {
        _logger.Dispose();

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
}
