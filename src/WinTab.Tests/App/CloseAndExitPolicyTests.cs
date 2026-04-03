using System.Reflection;
using FluentAssertions;
using WinTab.Core.Models;
using Xunit;

namespace WinTab.Tests.App;

public sealed class CloseAndExitPolicyTests
{
    [Fact]
    public void MainWindowClosePolicy_WhenTrayIconDisabled_ShouldNotMinimizeToTray()
    {
        MethodInfo? method = typeof(WinTab.App.Views.MainWindow).GetMethod(
            "ShouldMinimizeToTrayOnClose",
            BindingFlags.NonPublic | BindingFlags.Static);

        method.Should().NotBeNull("close policy must be explicitly testable");

        object? result = method?.Invoke(null, [false, false]);
        result.Should().BeOfType<bool>();
        ((bool)result!).Should().BeFalse(
            "without tray icon, closing the only window must not hide the app process");
    }

    [Fact]
    public void ExplorerOpenVerbExitPolicy_WhenInterceptionEnabled_ShouldAlwaysRestoreOnExit()
    {
        MethodInfo? method = typeof(WinTab.App.App).GetMethod(
            "ShouldDisableExplorerOpenVerbInterceptionOnExit",
            BindingFlags.NonPublic | BindingFlags.Static);

        method.Should().NotBeNull("exit policy should be centralized and testable");

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            PersistExplorerOpenVerbInterceptionAcrossExit = true
        };

        object? result = method?.Invoke(null, [settings]);
        result.Should().BeOfType<bool>();
        ((bool)result!).Should().BeTrue(
            "WinTab must stop hijacking Explorer once the process is no longer running");
    }

    [Fact]
    public void ExplorerOpenVerbExitPolicy_WhenPersistAcrossExitDisabled_ShouldRestoreOnExit()
    {
        MethodInfo? method = typeof(WinTab.App.App).GetMethod(
            "ShouldDisableExplorerOpenVerbInterceptionOnExit",
            BindingFlags.NonPublic | BindingFlags.Static);

        method.Should().NotBeNull("exit policy should be centralized and testable");

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            PersistExplorerOpenVerbInterceptionAcrossExit = false
        };

        object? result = method?.Invoke(null, [settings]);
        result.Should().BeOfType<bool>();
        ((bool)result!).Should().BeTrue();
    }

    [Fact]
    public void ExplorerOpenVerbExitPolicy_WhenInterceptionDisabled_ShouldNotTouchExplorerStateOnExit()
    {
        MethodInfo? method = typeof(WinTab.App.App).GetMethod(
            "ShouldDisableExplorerOpenVerbInterceptionOnExit",
            BindingFlags.NonPublic | BindingFlags.Static);

        method.Should().NotBeNull("exit policy should be centralized and testable");

        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = false,
            PersistExplorerOpenVerbInterceptionAcrossExit = true
        };

        object? result = method?.Invoke(null, [settings]);
        result.Should().BeOfType<bool>();
        ((bool)result!).Should().BeFalse(
            "when interception is disabled, shutdown should not mutate Explorer open-verb state");
    }
}
