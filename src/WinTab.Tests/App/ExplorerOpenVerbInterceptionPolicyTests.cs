using FluentAssertions;
using WinTab.App.Services;
using WinTab.Core.Models;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ExplorerOpenVerbInterceptionPolicyTests
{
    [Fact]
    public void NormalizeForNativeCurrentDirectoryBehavior_WhenChildFolderNewTabDisabled_ShouldNotMutateInterceptionFlag()
    {
        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            EnableExplorerOpenVerbInterception = false,
        };

        bool changed = ExplorerOpenVerbInterceptionPolicy.NormalizeForNativeCurrentDirectoryBehavior(settings);

        changed.Should().BeTrue();
        settings.EnableExplorerOpenVerbInterception.Should().BeTrue();
    }

    [Fact]
    public void NormalizeForNativeCurrentDirectoryBehavior_WhenAutoConvertDisabled_ShouldClearInterceptionFlag()
    {
        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = false,
            EnableExplorerOpenVerbInterception = true,
        };

        bool changed = ExplorerOpenVerbInterceptionPolicy.NormalizeForNativeCurrentDirectoryBehavior(settings);

        changed.Should().BeTrue();
        settings.EnableExplorerOpenVerbInterception.Should().BeFalse();
    }

    [Fact]
    public void ShouldEnableOpenVerbInterception_WhenAutoConvertEnabled_ShouldIgnoreChildFolderTabPreference()
    {
        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
            OpenChildFolderInNewTabFromActiveTab = false,
        };

        bool enabled = ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath: true);

        enabled.Should().BeTrue(
            "whether child folders stay in the current tab or open a new tab is a handler behavior choice, not a reason to disable the no-flicker interception transport");
    }

    [Fact]
    public void ShouldEnableOpenVerbInterception_WhenAutoConvertAndStablePathSatisfied_ShouldBeTrue()
    {
        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
        };

        bool enabled = ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath: true);

        enabled.Should().BeTrue();
    }

    [Fact]
    public void ShouldEnableOpenVerbInterception_WhenHandlerPathUnstable_ShouldBeFalse()
    {
        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = true,
        };

        bool enabled = ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath: false);

        enabled.Should().BeFalse();
    }

    [Fact]
    public void ShouldEnableOpenVerbInterception_WhenAutoConvertDisabled_ShouldBeFalse()
    {
        var settings = new AppSettings
        {
            EnableAutoConvertExplorerWindows = false,
            EnableExplorerOpenVerbInterception = true,
        };

        bool enabled = ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath: true);

        enabled.Should().BeFalse();
    }

    [Fact]
    public void ShouldPersistAcrossReboot_WhenRunAtStartupDisabledAndPersistDisabled_ShouldBeFalse()
    {
        var settings = new AppSettings
        {
            RunAtStartup = false,
            PersistExplorerOpenVerbInterceptionAcrossExit = false,
        };

        bool persist = ExplorerOpenVerbInterceptionPolicy.ShouldPersistAcrossReboot(settings);

        persist.Should().BeFalse();
    }

    [Fact]
    public void ShouldPersistAcrossReboot_WhenRunAtStartupEnabled_ShouldStillBeFalse()
    {
        var settings = new AppSettings
        {
            RunAtStartup = true,
            PersistExplorerOpenVerbInterceptionAcrossExit = false,
        };

        bool persist = ExplorerOpenVerbInterceptionPolicy.ShouldPersistAcrossReboot(settings);

        persist.Should().BeFalse(
            "WinTab must not leave Explorer hijacked after reboot just because auto-start is enabled");
    }

    [Fact]
    public void ShouldPersistAcrossReboot_WhenExplicitPersistEnabled_ShouldStillBeFalse()
    {
        var settings = new AppSettings
        {
            RunAtStartup = false,
            PersistExplorerOpenVerbInterceptionAcrossExit = true,
        };

        bool persist = ExplorerOpenVerbInterceptionPolicy.ShouldPersistAcrossReboot(settings);

        persist.Should().BeFalse(
            "WinTab must restore native Explorer behavior whenever the process is not running");
    }
}
