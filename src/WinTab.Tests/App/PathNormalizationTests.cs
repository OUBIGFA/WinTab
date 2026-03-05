using System;
using System.Reflection;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class PathNormalizationTests
{
    [Fact]
    public void TryNormalizeOpenTarget_WhenBareGuid_ShouldNormalizeToShellNamespace()
    {
        const string bareGuid = "{645FF040-5081-101B-9F08-00AA002F954E}";

        bool ok = InvokeTryNormalizeOpenTarget(bareGuid, out string normalized);

        ok.Should().BeTrue();
        Assert.Equal("::" + bareGuid, normalized);
    }

    private static bool InvokeTryNormalizeOpenTarget(string candidatePath, out string normalizedPath)
    {
        Type? type = Type.GetType("WinTab.ShellBridge.PathNormalization, WinTab.ShellBridge", throwOnError: false);
        type.Should().NotBeNull("PathNormalization type must be available through WinTab.ShellBridge assembly");

        MethodInfo? method = type!.GetMethod(
            "TryNormalizeOpenTarget",
            BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
        method.Should().NotBeNull("TryNormalizeOpenTarget must exist");

        object?[] args = [candidatePath, null];
        bool result = (bool)(method!.Invoke(null, args) ?? false);
        normalizedPath = args[1] as string ?? string.Empty;
        return result;
    }
}
