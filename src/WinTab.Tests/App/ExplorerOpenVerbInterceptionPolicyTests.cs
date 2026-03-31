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
    public void ShouldEnableOpenVerbInterception_WhenChildFolderNewTabDisabled_ShouldRemainTrueForDirectReuse()
    {
        var settings = new AppSettings
        {
            OpenChildFolderInNewTabFromActiveTab = false,
            EnableExplorerOpenVerbInterception = true,
        };

        bool enabled = ExplorerOpenVerbInterceptionPolicy.ShouldEnableOpenVerbInterception(settings, hasStableOpenVerbHandlerPath: true);

        enabled.Should().BeTrue();
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
    public void ShouldPersistAcrossReboot_WhenRunAtStartupEnabled_ShouldBeTrue()
    {
        var settings = new AppSettings
        {
            RunAtStartup = true,
            PersistExplorerOpenVerbInterceptionAcrossExit = false,
        };

        bool persist = ExplorerOpenVerbInterceptionPolicy.ShouldPersistAcrossReboot(settings);

        persist.Should().BeTrue();
    }

    [Fact]
    public void ShouldPersistAcrossReboot_WhenExplicitPersistEnabled_ShouldBeTrue()
    {
        var settings = new AppSettings
        {
            RunAtStartup = false,
            PersistExplorerOpenVerbInterceptionAcrossExit = true,
        };

        bool persist = ExplorerOpenVerbInterceptionPolicy.ShouldPersistAcrossReboot(settings);

        persist.Should().BeTrue();
    }
}
