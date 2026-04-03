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
    [InlineData("shell:RecycleBinFolder", false, "tab-navigable shell namespace must merge into tab")]
    [InlineData("shell::Downloads", false, "tab-navigable shell namespace must merge into tab")]
    [InlineData("::{645FF040-5081-101B-9F08-00AA002F954E}", false, "Recycle Bin must merge into tab")]
    [InlineData("{645FF040-5081-101B-9F08-00AA002F954E}", false, "GUID shell namespace must merge into tab")]
    [InlineData("C:\\Windows", false, "physical path must merge into tab")]
    [InlineData("\\\\server\\share", false, "UNC path must merge into tab")]
    public void ShouldBypassAutoConvertForLocation_ShouldOnlyBypassNativeShellNamespaces(string input, bool expected, string reason)
    {
        bool actual = InvokePrivateStatic<bool>("ShouldBypassAutoConvertForLocation", input);
        actual.Should().Be(expected, reason);
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData(@"C:\Users\bigfa\Desktop", true)]
    [InlineData("shell:RecycleBinFolder", true, "Recycle Bin must convert to tab")]
    [InlineData("::{645FF040-5081-101B-9F08-00AA002F954E}", true, "Recycle Bin GUID must convert to tab")]
    [InlineData("shell::Downloads", true, "Downloads shell namespace must convert to tab")]
    public void ShouldConvertWindowLocationToTab_ShouldAllowShellNamespacesToConvert(string? input, bool expected, string reason = "")
    {
        bool actual = InvokePrivateStatic<bool>("ShouldConvertWindowLocationToTab", input);
        actual.Should().Be(expected, reason);
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

    [Fact]
    public void SelectBestExplorerTargetCandidate_WhenOnlyMinimizedWindowAndFallbackEnabled_ShouldReuseMinimizedWindow()
    {
        IntPtr minimized = new(0x216D6);
        (IntPtr Hwnd, bool IsMinimized, int TabCount)[] candidates =
        [
            (minimized, true, 1)
        ];

        IntPtr picked = InvokePrivateStatic<IntPtr>(
            "SelectBestExplorerTargetCandidate",
            candidates,
            IntPtr.Zero,
            minimized,
            true);

        picked.Should().Be(minimized,
            "when all Explorer windows are minimized, interception should still pick one for reuse instead of forcing standalone launch");
    }

    [Fact]
    public void SelectBestExplorerTargetCandidate_WhenOnlyMinimizedWindowAndFallbackDisabled_ShouldReturnZero()
    {
        IntPtr minimized = new(0x216D6);
        (IntPtr Hwnd, bool IsMinimized, int TabCount)[] candidates =
        [
            (minimized, true, 1)
        ];

        IntPtr picked = InvokePrivateStatic<IntPtr>(
            "SelectBestExplorerTargetCandidate",
            candidates,
            IntPtr.Zero,
            minimized,
            false);

        picked.Should().Be(IntPtr.Zero);
    }

    [Fact]
    public void SelectBestExplorerTargetCandidate_WhenNonMinimizedCandidateExists_ShouldPreferItEvenWithFallbackEnabled()
    {
        IntPtr minimized = new(0x216D6);
        IntPtr normal = new(0x40240);

        (IntPtr Hwnd, bool IsMinimized, int TabCount)[] candidates =
        [
            (minimized, true, 6),
            (normal, false, 1)
        ];

        IntPtr picked = InvokePrivateStatic<IntPtr>(
            "SelectBestExplorerTargetCandidate",
            candidates,
            IntPtr.Zero,
            minimized,
            true);

        picked.Should().Be(normal,
            "minimized fallback should only kick in when no non-minimized reusable Explorer target exists");
    }

    [Fact]
    public void BuildReusableExplorerCandidates_WhenOnlyInvisibleMinimizedWindowAndFallbackEnabled_ShouldKeepCandidate()
    {
        IntPtr minimized = new(0x3104C);
        (IntPtr Hwnd, bool IsVisible, bool IsMinimized, int TabCount)[] windows =
        [
            (minimized, false, true, 1)
        ];

        List<(IntPtr Hwnd, bool IsMinimized, int TabCount)> candidates = InvokePrivateStatic<List<(IntPtr Hwnd, bool IsMinimized, int TabCount)>>(
            "BuildReusableExplorerCandidates",
            windows,
            IntPtr.Zero,
            true);

        candidates.Should().ContainSingle();
        candidates[0].Hwnd.Should().Be(minimized,
            "minimized Explorer windows can be invisible in EnumWindows results but still must remain eligible for reuse fallback");
    }

    [Fact]
    public void BuildReusableExplorerCandidates_WhenOnlyInvisibleMinimizedWindowAndFallbackDisabled_ShouldDropCandidate()
    {
        IntPtr minimized = new(0x3104C);
        (IntPtr Hwnd, bool IsVisible, bool IsMinimized, int TabCount)[] windows =
        [
            (minimized, false, true, 1)
        ];

        List<(IntPtr Hwnd, bool IsMinimized, int TabCount)> candidates = InvokePrivateStatic<List<(IntPtr Hwnd, bool IsMinimized, int TabCount)>>(
            "BuildReusableExplorerCandidates",
            windows,
            IntPtr.Zero,
            false);

        candidates.Should().BeEmpty();
    }

    [Fact]
    public void BuildReusableExplorerCandidates_WhenInvisibleNonMinimizedWindowAndFallbackEnabled_ShouldStillDropCandidate()
    {
        IntPtr hidden = new(0x40240);
        (IntPtr Hwnd, bool IsVisible, bool IsMinimized, int TabCount)[] windows =
        [
            (hidden, false, false, 2)
        ];

        List<(IntPtr Hwnd, bool IsMinimized, int TabCount)> candidates = InvokePrivateStatic<List<(IntPtr Hwnd, bool IsMinimized, int TabCount)>>(
            "BuildReusableExplorerCandidates",
            windows,
            IntPtr.Zero,
            true);

        candidates.Should().BeEmpty(
            "fallback should reuse minimized Explorer windows, not arbitrary hidden windows");
    }

    [Fact]
    public void RuntimeHookSource_ShouldRestoreTrackedExplorerWindowsOnDispose()
    {
        string source = File.ReadAllText(TestRepoPaths.GetFile(["src", "WinTab.App", "ExplorerTabUtilityPort", "ExplorerTabHookService.cs"]));

        source.Should().Contain("RestoreTrackedExplorerWindowsOnShutdown();",
            "disposing the runtime hook must unhide any Explorer windows that WinTab hid during conversion");
        source.Should().Contain(".. _earlyHiddenExplorer.Keys",
            "shutdown restore must include Explorer windows hidden by the early anti-flash path");
        source.Should().Contain(".. _pending.Keys",
            "shutdown restore must include Explorer windows hidden while a conversion was still in flight");
        source.Should().Contain("_windowManager.Show(hwnd);",
            "shutdown restore must explicitly make tracked Explorer windows visible again");
    }

    private static T InvokePrivateStatic<T>(string methodName, params object?[] args)
    {
        MethodInfo method = typeof(ExplorerTabHookService).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException($"Method not found: {methodName}");

        object? result = method.Invoke(null, args);
        return result is T typed ? typed : throw new InvalidOperationException($"Unexpected return type for {methodName}.");
    }
}
