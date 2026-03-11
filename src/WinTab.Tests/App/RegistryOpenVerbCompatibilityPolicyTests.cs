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
    public void ShouldPreferDelegateExecuteOverride_WhenComHostExists_ShouldRemainDisabledForCompatibility()
    {
        var interceptor = new RegistryOpenVerbInterceptor(@"C:\Program Files\WinTab\WinTab.exe", _logger);

        bool preferDelegateExecute = InvokeShouldPreferDelegateExecuteOverride(interceptor, comHostExists: true);

        preferDelegateExecute.Should().BeFalse(
            "the open verb override must stay command-based so 32-bit third-party shell hosts do not fail with COM class registration errors");
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

    private static bool InvokeShouldPreferDelegateExecuteOverride(RegistryOpenVerbInterceptor interceptor, bool comHostExists)
    {
        MethodInfo method = typeof(RegistryOpenVerbInterceptor).GetMethod(
            "ShouldPreferDelegateExecuteOverride",
            BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Method not found: ShouldPreferDelegateExecuteOverride");

        object? result = method.Invoke(interceptor, [comHostExists]);
        return result is bool value && value;
    }
}
