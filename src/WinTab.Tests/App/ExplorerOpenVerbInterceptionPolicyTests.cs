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
            OpenChildFolderInNewTabFromActiveTab = false,
            EnableExplorerOpenVerbInterception = true,
        };

        bool changed = ExplorerOpenVerbInterceptionPolicy.NormalizeForNativeCurrentDirectoryBehavior(settings);

        changed.Should().BeFalse();
        settings.EnableExplorerOpenVerbInterception.Should().BeTrue();
    }

    [Fact]
    public void NormalizeForNativeCurrentDirectoryBehavior_WhenChildFolderNewTabEnabled_ShouldKeepInterceptionFlag()
    {
        var settings = new AppSettings
        {
            OpenChildFolderInNewTabFromActiveTab = true,
            EnableExplorerOpenVerbInterception = false,
        };

        bool changed = ExplorerOpenVerbInterceptionPolicy.NormalizeForNativeCurrentDirectoryBehavior(settings);

        changed.Should().BeFalse();
        settings.EnableExplorerOpenVerbInterception.Should().BeFalse();
    }

    [Fact]
    public void ShouldEnableOpenVerbInterception_WhenChildFolderNewTabDisabled_ShouldBeFalseForNativeBrowsing()
    {
        var settings = new AppSettings
        {
            OpenChildFolderInNewTabFromActiveTab = false,
            EnableExplorerOpenVerbInterception = true,
        };

        bool enabled = ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath: true);

        enabled.Should().BeFalse(
            "when child folders should open with native in-place browsing, WinTab must not register the open-verb interceptor");
    }

    [Fact]
    public void ShouldEnableOpenVerbInterception_WhenAllConditionsSatisfied_ShouldBeTrue()
    {
        var settings = new AppSettings
        {
            OpenChildFolderInNewTabFromActiveTab = true,
            EnableExplorerOpenVerbInterception = true,
        };

        bool enabled = ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath: true);

        enabled.Should().BeTrue();
    }

    [Fact]
    public void ShouldEnableOpenVerbInterception_WhenHandlerPathUnstable_ShouldBeFalse()
    {
        var settings = new AppSettings
        {
            OpenChildFolderInNewTabFromActiveTab = true,
            EnableExplorerOpenVerbInterception = true,
        };

        bool enabled = ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath: false);

        enabled.Should().BeFalse();
    }

    [Fact]
    public void ShouldEnableOpenVerbInterception_WhenSettingDisabled_ShouldBeFalse()
    {
        var settings = new AppSettings
        {
            OpenChildFolderInNewTabFromActiveTab = true,
            EnableExplorerOpenVerbInterception = false,
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
