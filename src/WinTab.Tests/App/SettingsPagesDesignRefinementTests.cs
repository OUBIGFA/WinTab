using System.IO;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public class SettingsPagesDesignRefinementTests
{
    [Fact]
    public void GeneralAndBehaviorPages_ShouldKeepRefinedLayoutMarkers()
    {
        string generalPagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "GeneralPage.xaml");
        string behaviorPagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "BehaviorPage.xaml");

        XDocument generalPage = XDocument.Load(generalPagePath);
        XDocument behaviorPage = XDocument.Load(behaviorPagePath);

        string? generalMaxWidth = generalPage
            .Descendants()
            .First(e => e.Name.LocalName == "StackPanel" && e.Attribute("MaxWidth") is not null)
            .Attribute("MaxWidth")?.Value;

        string? behaviorMaxWidth = behaviorPage
            .Descendants()
            .First(e => e.Name.LocalName == "StackPanel" && e.Attribute("MaxWidth") is not null)
            .Attribute("MaxWidth")?.Value;

        generalMaxWidth.Should().Be("880", "General page should use the wider polished settings rhythm");
        behaviorMaxWidth.Should().Be("880", "Behavior page should use the wider polished settings rhythm");

        var generalRows = generalPage
            .Descendants()
            .Where(e => e.Name.LocalName == "Border" && (string?)e.Attribute("Style") == "{StaticResource SettingRowBorderStyle}")
            .ToList();
        var behaviorRows = behaviorPage
            .Descendants()
            .Where(e => e.Name.LocalName == "Border" && (string?)e.Attribute("Style") == "{StaticResource SettingRowBorderStyle}")
            .ToList();

        generalRows.Should().HaveCount(5, "General page should render five dedicated setting rows");
        behaviorRows.Should().HaveCount(4, "Behavior page should render four dedicated setting rows");

        foreach (XElement row in generalRows)
        {
            row.Attribute("Style")?.Value.Should().Be("{StaticResource SettingRowBorderStyle}", "General rows should share the dedicated setting row style");
        }

        foreach (XElement row in behaviorRows)
        {
            row.Attribute("Style")?.Value.Should().Be("{StaticResource SettingRowBorderStyle}", "Behavior rows should share the dedicated setting row style");
        }
    }

    [Fact]
    public void BehaviorPage_ShouldPlaceDoubleClickCloseRuleFirst()
    {
        string behaviorPagePath = GetProjectFilePath("WinTab.App", "Views", "Pages", "BehaviorPage.xaml");
        XDocument behaviorPage = XDocument.Load(behaviorPagePath);

        var ruleTitles = behaviorPage
            .Descendants()
            .Where(e => e.Name.LocalName == "TextBlock" &&
                        (string?)e.Attribute("Style") == "{StaticResource SettingTitleStyle}")
            .Select(e => (string?)e.Attribute("Text"))
            .Where(v => v is not null && v.StartsWith("{DynamicResource Behavior_"))
            .Skip(1)
            .Take(3)
            .ToList();

        ruleTitles.Should().ContainInOrder(
            [
                "{DynamicResource Behavior_CloseTabOnDoubleClick}",
                "{DynamicResource Behavior_OpenNewTabFromActiveTabPath}",
                "{DynamicResource Behavior_OpenChildFolderInNewTab}"
            ],
            "tab rules should surface the most direct close action first");
    }

    private static string GetProjectFilePath(string projectFolder, params string[] parts)
    {
        string[] allParts = ["src", projectFolder, .. parts];
        return TestRepoPaths.GetFile(allParts);
    }
}
