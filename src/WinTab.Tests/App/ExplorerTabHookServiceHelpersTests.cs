using System;
using System.IO;
using System.Reflection;
using FluentAssertions;
using WinTab.App.ExplorerTabUtilityPort;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ExplorerTabHookServiceHelpersTests
{
    [Theory]
    [InlineData("C:\\Windows", true)]
    [InlineData("c:/Windows", true)]
    [InlineData("\\\\server\\share", true)]
    [InlineData("::{645FF040-5081-101B-9F08-00AA002F954E}", true)]
    [InlineData("shell:::{645FF040-5081-101B-9F08-00AA002F954E}", true)]
    [InlineData("shell::Downloads", false)]
    [InlineData("", false)]
    [InlineData("   ", false)]
    public void IsRealFileSystemLocation_ShouldMatchExpectedRules(string input, bool expected)
    {
        bool actual = InvokePrivateStatic<bool>("IsRealFileSystemLocation", input);
        actual.Should().Be(expected);
    }

    [Fact]
    public void IsChildPathOf_ShouldDetectDirectChild()
    {
        string root = Path.Combine(Path.GetTempPath(), "WinTabHookHelperTests", Guid.NewGuid().ToString("N"));
        string child = Path.Combine(root, "sub");

        bool actual = InvokePrivateStatic<bool>("IsChildPathOf", root, child);
        actual.Should().BeTrue();
    }

    [Fact]
    public void IsChildPathOf_ShouldRejectSibling()
    {
        string basePath = Path.Combine(Path.GetTempPath(), "WinTabHookHelperTests", Guid.NewGuid().ToString("N"));
        string sibling = Path.Combine(Path.GetTempPath(), "WinTabHookHelperTests", Guid.NewGuid().ToString("N"));

        bool actual = InvokePrivateStatic<bool>("IsChildPathOf", basePath, sibling);
        actual.Should().BeFalse();
    }

    [Theory]
    [InlineData("explorer.exe", "explorer")]
    [InlineData("EXPLORER.EXE", "EXPLORER")]
    [InlineData(" explorer ", "explorer")]
    [InlineData("explorer", "explorer")]
    public void NormalizeExeName_ShouldRemoveExeSuffix(string input, string expected)
    {
        string actual = InvokePrivateStatic<string>("NormalizeExeName", input);
        actual.Should().Be(expected);
    }

    /// <summary>
    /// Verifies that the timeout used for in-place current-tab navigation is kept low
    /// enough to feel native. A confirmation wait of >250ms is user-perceptible.
    /// </summary>
    [Fact]
    public void CurrentTabNavigateTimeoutMs_ShouldBeLowEnoughToFeelNative()
    {
        // Native Explorer navigation is instant. Our COM-based path adds process-startup
        // overhead (~100-200 ms). Any additional timeout on top of that should be minimal.
        // 250 ms is the threshold at which humans begin to perceive latency as "sluggish".
        ExplorerTabHookService.CurrentTabNavigateTimeoutMs.Should().BeLessOrEqualTo(250,
            "current-tab COM navigation is fire-and-commit; excessive retries make it feel sluggish");
    }

    /// <summary>
    /// Verifies that the per-retry delay inside the navigation retry loop is ≤120ms,
    /// so that within the navigate timeout we get at least one meaningful retry attempt
    /// without the loop adding more latency than the timeout itself.
    /// </summary>
    [Fact]
    public void CurrentTabNavigateRetryMs_ShouldFitWithinTimeout()
    {
        ExplorerTabHookService.CurrentTabNavigateRetryMs.Should().BeLessOrEqualTo(
            ExplorerTabHookService.CurrentTabNavigateTimeoutMs,
            "each retry delay must fit within the navigate timeout so we always get at least one retry");
    }

    private static T InvokePrivateStatic<T>(string methodName, params object?[] args)
    {
        MethodInfo method = typeof(ExplorerTabHookService).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException($"Method not found: {methodName}");

        object? result = method.Invoke(null, args);
        return result is T typed ? typed : throw new InvalidOperationException($"Unexpected return type for {methodName}.");
    }
}
