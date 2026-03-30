using System.IO;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class SettingsPagesLayoutRegressionTests
{
    [Fact]
    public void GeneralAndBehaviorPages_ShouldUseDedicatedSettingRowsInsteadOfCardControlHeaders()
    {
        string generalPagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "GeneralPage.xaml");
        string behaviorPagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "BehaviorPage.xaml");

        XDocument generalPage = XDocument.Load(generalPagePath);
        XDocument behaviorPage = XDocument.Load(behaviorPagePath);

        generalPage.Descendants().Count(e => e.Name.LocalName == "CardControl").Should().Be(0,
            "the startup page labels should not depend on CardControl header rendering");
        behaviorPage.Descendants().Count(e => e.Name.LocalName == "CardControl").Should().Be(0,
            "the tab rules page labels should not depend on CardControl header rendering");
    }

    private static string GetProjectFilePath(string projectFolder, params string[] parts)
    {
        string[] allParts = ["src", projectFolder, .. parts];
        return TestRepoPaths.GetFile(allParts);
    }
}
