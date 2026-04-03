using FluentAssertions;
using WinTab.Core.Models;
using Xunit;

namespace WinTab.Tests.App;

public sealed class AppSettingsDefaultsTests
{
    [Fact]
    public void FreshInstallDefaults_ShouldKeepExplorerHooksDisabled()
    {
        var settings = new AppSettings();

        settings.EnableExplorerOpenVerbInterception.Should().BeFalse(
            "fresh installs must not take over Explorer open verbs until the user explicitly enables the feature");
        settings.EnableAutoConvertExplorerWindows.Should().BeFalse(
            "fresh installs must not start auto-converting Explorer windows before the user opts in");
    }
}
