using System;
using System.IO;
using System.Reflection;
using FluentAssertions;
using WinTab.App.Services;
using WinTab.Diagnostics;
using Xunit;

namespace WinTab.Tests.App;

public sealed class RegistryOpenVerbCompatibilityPolicyTests : IDisposable
{
    private readonly string _tempDir;
    private readonly Logger _logger;

    public RegistryOpenVerbCompatibilityPolicyTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "WinTabRegistryPolicyTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _logger = new Logger(Path.Combine(_tempDir, "policy-tests.log"));
    }

    [Fact]
    public void ShouldPreferDelegateExecuteOverride_WhenBothBridgeArchitecturesExist_ShouldPreferNativeExplorerBridge()
    {
        var interceptor = new RegistryOpenVerbInterceptor(@"C:\Program Files\WinTab\WinTab.exe", _logger);

        bool preferDelegateExecute = InvokeShouldPreferDelegateExecuteOverride(
            interceptor,
            comHostExists: true,
            comHost32Exists: true,
            x64RuntimeCompatible: true,
            x86RuntimeCompatible: true);

        preferDelegateExecute.Should().BeTrue(
            "when both COM bridge architectures are available, Explorer browsing should stay on the DelegateExecute bridge without breaking 32-bit shell hosts");
    }

    [Fact]
    public void ShouldPreferDelegateExecuteOverride_When32BitBridgeIsMissing_ShouldStayInLegacyMode()
    {
        var interceptor = new RegistryOpenVerbInterceptor(@"C:\Program Files\WinTab\WinTab.exe", _logger);

        bool preferDelegateExecute = InvokeShouldPreferDelegateExecuteOverride(
            interceptor,
            comHostExists: true,
            comHost32Exists: false,
            x64RuntimeCompatible: true,
            x86RuntimeCompatible: true);

        preferDelegateExecute.Should().BeFalse(
            "global DelegateExecute registration is only safe when both 64-bit Explorer and 32-bit third-party shell hosts can load a matching in-proc bridge");
    }

    [Fact]
    public void ShouldPreferDelegateExecuteOverride_When32BitRuntimeIsMissing_ShouldStayInLegacyMode()
    {
        var interceptor = new RegistryOpenVerbInterceptor(@"C:\Program Files\WinTab\WinTab.exe", _logger);

        bool preferDelegateExecute = InvokeShouldPreferDelegateExecuteOverride(
            interceptor,
            comHostExists: true,
            comHost32Exists: true,
            x64RuntimeCompatible: true,
            x86RuntimeCompatible: false);

        preferDelegateExecute.Should().BeFalse(
            "global DelegateExecute registration is only safe when the 32-bit shell host can activate a compatible bridge runtime instead of failing with class-not-registered");
    }

    [Fact]
    public void ShouldPreferDelegateExecuteOverride_WhenBridgePrerequisitesAreUnavailable_ShouldStayInLegacyMode()
    {
        var interceptor = new RegistryOpenVerbInterceptor(@"C:\Program Files\WinTab\WinTab.exe", _logger);

        bool preferDelegateExecute = InvokeShouldPreferDelegateExecuteOverride(
            interceptor,
            comHostExists: true,
            comHost32Exists: true,
            x64RuntimeCompatible: false,
            x86RuntimeCompatible: true);

        preferDelegateExecute.Should().BeFalse(
            "runtime-only interception must fall back to command overrides whenever the DelegateExecute bridge is incomplete");
    }

    [Fact]
    public void ShouldPreferDelegateExecuteOverride_WhenBridgePrerequisitesAreReady_ShouldKeepDelegateExecuteMode()
    {
        var interceptor = new RegistryOpenVerbInterceptor(@"C:\Program Files\WinTab\WinTab.exe", _logger);

        bool preferDelegateExecute = InvokeShouldPreferDelegateExecuteOverride(
            interceptor,
            comHostExists: true,
            comHost32Exists: true,
            x64RuntimeCompatible: true,
            x86RuntimeCompatible: true);

        preferDelegateExecute.Should().BeTrue(
            "runtime-only interception should still prefer the no-flicker DelegateExecute path when both bridge architectures are usable");
    }

    [Fact]
    public void ShouldResetOverrideBeforeRepair_WhenIncompleteOverrideExistsWithoutBackup_ShouldResetBeforeTakingNewBackup()
    {
        bool shouldReset = InvokeShouldResetOverrideBeforeRepair(
            registryPointsToUs: false,
            registryPointsToAnyWinTabHandler: true,
            hasAnyBackup: false);

        shouldReset.Should().BeTrue(
            "an incomplete HKCU-only WinTab override must not be captured as the new baseline backup");
    }

    [Fact]
    public void ShouldResetOverrideBeforeRepair_WhenBackupAlreadyExists_ShouldNotResetSolelyForRepair()
    {
        bool shouldReset = InvokeShouldResetOverrideBeforeRepair(
            registryPointsToUs: false,
            registryPointsToAnyWinTabHandler: true,
            hasAnyBackup: true);

        shouldReset.Should().BeFalse(
            "once a valid backup exists, repair can use it instead of forcing a pre-emptive reset");
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

    private static bool InvokeShouldPreferDelegateExecuteOverride(
        RegistryOpenVerbInterceptor interceptor,
        bool comHostExists,
        bool comHost32Exists,
        bool x64RuntimeCompatible,
        bool x86RuntimeCompatible)
    {
        MethodInfo method = typeof(RegistryOpenVerbInterceptor).GetMethod(
            "ShouldPreferDelegateExecuteOverride",
            BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Method not found: ShouldPreferDelegateExecuteOverride");

        object? result = method.Invoke(interceptor, [comHostExists, comHost32Exists, x64RuntimeCompatible, x86RuntimeCompatible]);
        return result is bool value && value;
    }

    private static bool InvokeShouldResetOverrideBeforeRepair(
        bool registryPointsToUs,
        bool registryPointsToAnyWinTabHandler,
        bool hasAnyBackup)
    {
        MethodInfo method = typeof(RegistryOpenVerbInterceptor).GetMethod(
            "ShouldResetOverrideBeforeRepair",
            BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Method not found: ShouldResetOverrideBeforeRepair");

        object? result = method.Invoke(null, [registryPointsToUs, registryPointsToAnyWinTabHandler, hasAnyBackup]);
        return result is bool value && value;
    }
}
