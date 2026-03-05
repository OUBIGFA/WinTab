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
            EnableExplorerOpenVerbInterception = true,
        };

        bool changed = ExplorerOpenVerbInterceptionPolicy.NormalizeForNativeCurrentDirectoryBehavior(settings);

        changed.Should().BeFalse();
        settings.EnableExplorerOpenVerbInterception.Should().BeTrue();
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
}
