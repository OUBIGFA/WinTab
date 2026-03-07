using System;
using System.Reflection;
using FluentAssertions;
using WinTab.ShellBridge;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ShellBridgePipeLatencyBudgetTests
{
    [Fact]
    public void DelegateExecutePipeFallbackBudget_ShouldBeLowEnoughToAvoidTaskbarFreeze()
    {
        Type pipeClientType = typeof(WinTabOpenFolderDelegateExecute).Assembly.GetType("WinTab.ShellBridge.OpenRequestPipeClient")
            ?? throw new InvalidOperationException("OpenRequestPipeClient type not found.");

        int defaultConnectTimeoutMs = ReadPrivateConstInt(pipeClientType, "DefaultConnectTimeoutMs");
        int retryConnectTimeoutMs = ReadPrivateConstInt(pipeClientType, "RetryConnectTimeoutMs");
        int retryDelayMs = ReadPrivateConstInt(pipeClientType, "RetryDelayMs");

        int totalBlockingBudgetMs = defaultConnectTimeoutMs + retryConnectTimeoutMs + retryDelayMs;

        totalBlockingBudgetMs.Should().BeLessOrEqualTo(300,
            "DelegateExecute runs inside explorer.exe. Excessive sync pipe wait freezes taskbar interactions.");
        defaultConnectTimeoutMs.Should().BeLessOrEqualTo(150);
        retryConnectTimeoutMs.Should().BeLessOrEqualTo(150);
        retryDelayMs.Should().BeLessOrEqualTo(50);
    }

    [Fact]
    public void IsTaskbarForegroundWindow_WhenResolverReturnsTaskbarClass_ShouldBeTrue()
    {
        Type delegateExecuteType = typeof(WinTabOpenFolderDelegateExecute);
        MethodInfo method = delegateExecuteType.GetMethod(
            "IsTaskbarForegroundWindow",
            BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("IsTaskbarForegroundWindow method not found.");

        Func<nint, string> resolver = static _ => "Shell_TrayWnd";
        object? result = method.Invoke(null, [new nint(123), resolver]);

        (result as bool?).Should().BeTrue();
    }

    [Fact]
    public void IsTaskbarForegroundWindow_WhenResolverReturnsOtherClass_ShouldBeFalse()
    {
        Type delegateExecuteType = typeof(WinTabOpenFolderDelegateExecute);
        MethodInfo method = delegateExecuteType.GetMethod(
            "IsTaskbarForegroundWindow",
            BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("IsTaskbarForegroundWindow method not found.");

        Func<nint, string> resolver = static _ => "CabinetWClass";
        object? result = method.Invoke(null, [new nint(123), resolver]);

        (result as bool?).Should().BeFalse();
    }

    private static int ReadPrivateConstInt(Type type, string fieldName)
    {
        FieldInfo field = type.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException($"Field not found: {fieldName}");

        object? value = field.GetRawConstantValue();
        return value is int intValue
            ? intValue
            : throw new InvalidOperationException($"Unexpected value for field: {fieldName}");
    }
}
