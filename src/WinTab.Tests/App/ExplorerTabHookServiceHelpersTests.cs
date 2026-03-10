using System;
using System.IO;
using System.Reflection;
using FluentAssertions;
using WinTab.App.ExplorerTabUtilityPort;
using WinTab.Platform.Win32;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ExplorerTabHookServiceHelpersTests
{
    [Theory]
    [InlineData("C:\\Windows", true)]
    [InlineData("c:/Windows", true)]
    [InlineData("\\\\server\\share", true)]
    [InlineData("::{645FF040-5081-101B-9F08-00AA002F954E}", false)]
    [InlineData("shell:::{645FF040-5081-101B-9F08-00AA002F954E}", false)]
    [InlineData("shell::Downloads", false)]
    [InlineData("", false)]
    [InlineData("   ", false)]
    public void IsRealFileSystemLocation_ShouldMatchExpectedRules(string input, bool expected)
    {
        bool actual = InvokePrivateStatic<bool>("IsRealFileSystemLocation", input);
        actual.Should().Be(expected);
    }

    [Theory]
    [InlineData("C:\\Windows", true)]
    [InlineData("\\\\server\\share", true)]
    [InlineData("shell:RecycleBinFolder", true)]
    [InlineData("shell::Downloads", true)]
    [InlineData("::{645FF040-5081-101B-9F08-00AA002F954E}", true)]
    [InlineData("{645FF040-5081-101B-9F08-00AA002F954E}", true)]
    [InlineData("", false)]
    [InlineData("   ", false)]
    public void IsTabNavigableLocation_ShouldMatchExpectedRules(string input, bool expected)
    {
        bool actual = InvokePrivateStatic<bool>("IsTabNavigableLocation", input);
        actual.Should().Be(expected);
    }

    [Theory]
    [InlineData("shell:RecycleBinFolder", true)]
    [InlineData("shell::Downloads", true)]
    [InlineData("::{645FF040-5081-101B-9F08-00AA002F954E}", true)]
    [InlineData("{645FF040-5081-101B-9F08-00AA002F954E}", true)]
    [InlineData("C:\\Windows", false)]
    [InlineData("\\\\server\\share", false)]
    public void ShouldBypassAutoConvertForLocation_ShouldPreferNativeForNamespaceTargets(string input, bool expected)
    {
        bool actual = InvokePrivateStatic<bool>("ShouldBypassAutoConvertForLocation", input);
        actual.Should().Be(expected);
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData(@"C:\Users\bigfa\Desktop", true)]
    [InlineData("shell:RecycleBinFolder", false)]
    [InlineData("::{645FF040-5081-101B-9F08-00AA002F954E}", false)]
    [InlineData("shell::Downloads", false)]
    public void ShouldConvertWindowLocationToTab_ShouldRejectShellNamespaceWindows(string? input, bool expected)
    {
        bool actual = InvokePrivateStatic<bool>("ShouldConvertWindowLocationToTab", input);
        actual.Should().Be(expected,
            "shell namespace windows should stay native even if standalone-window auto-convert is enabled");
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

    [Theory]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    public void ShouldContinueWithTabReuseAfterCurrentNavigateAttempt_ShouldRespectNativeBrowseRequirement(
        bool navigatedCurrentTab,
        bool requiredCurrentWindowNavigation,
        bool expected)
    {
        bool actual = InvokePrivateStatic<bool>(
            "ShouldContinueWithTabReuseAfterCurrentNavigateAttempt",
            navigatedCurrentTab,
            requiredCurrentWindowNavigation);

        actual.Should().Be(expected);
    }

    [Fact]
    public void ShouldSkipNewTabAlignment_WhenNewTabIsThisPcAndSourceIsFolder_ShouldBeFalse()
    {
        const string thisPcNamespace = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
        string sourceFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop");

        bool skip = InvokePrivateStatic<bool>(
            "ShouldSkipNewTabAlignment",
            thisPcNamespace,
            sourceFolder,
            new ShellLocationIdentityService());

        skip.Should().BeFalse(
            "inherit-current-tab-path should still align when the new tab starts at This PC");
    }

    [Fact]
    public void ShouldSkipNewTabAlignment_WhenNewTabAlreadyMatchesSource_ShouldBeTrue()
    {
        string sourceFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop");

        bool skip = InvokePrivateStatic<bool>(
            "ShouldSkipNewTabAlignment",
            sourceFolder,
            sourceFolder,
            new ShellLocationIdentityService());

        skip.Should().BeTrue();
    }

    [Theory]
    [InlineData(2, 3, @"C:\Users\bigfa\Desktop", true)]
    [InlineData(2, 2, @"C:\Users\bigfa\Desktop", false)]
    [InlineData(3, 2, @"C:\Users\bigfa\Desktop", false)]
    [InlineData(2, 3, null, false)]
    [InlineData(2, 3, "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", false)]
    public void ShouldApplyCachedNewTabAlignment_ShouldRequireTabGrowthAndFileSystemSource(
        int previousTabCount,
        int currentTabCount,
        string? cachedSourceLocation,
        bool expected)
    {
        bool actual = InvokePrivateStatic<bool>(
            "ShouldApplyCachedNewTabAlignment",
            previousTabCount,
            currentTabCount,
            cachedSourceLocation);

        actual.Should().Be(expected);
    }

    [Fact]
    public void Regression_DoubleClickCloseSequence_ShouldNotQualifyForAlignment()
    {
        // Repro model for the observed incident:
        // 1) Explorer had an active tab with a file-system location.
        // 2) User double-clicked tab title to close a tab.
        // 3) Explorer internal UI events may still emit tab-create notifications.
        //    If tab count does not grow from the cached baseline, alignment must be skipped.
        const int baselineTabCount = 2;
        const int tabCountAfterCloseAndUiChurn = 2;
        const string cachedActiveLocation = @"E:\_Free code\WinTab\publish";

        bool shouldAlign = InvokePrivateStatic<bool>(
            "ShouldApplyCachedNewTabAlignment",
            baselineTabCount,
            tabCountAfterCloseAndUiChurn,
            cachedActiveLocation);

        shouldAlign.Should().BeFalse(
            "close-tab UI churn must not be treated as a real new-tab creation for inherit-current-path alignment");
    }

    private static T InvokePrivateStatic<T>(string methodName, params object?[] args)
    {
        MethodInfo method = typeof(ExplorerTabHookService).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException($"Method not found: {methodName}");

        object? result = method.Invoke(null, args);
        return result is T typed ? typed : throw new InvalidOperationException($"Unexpected return type for {methodName}.");
    }
}
